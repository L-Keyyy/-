import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx";

const defaultFiles = ["W27", "W28", "W29", "W30"].map((week) => ({
  week,
  file: `/Users/gl001426/Desktop/临时/${week}.xlsx`,
}));
const args = process.argv.slice(2);
const files = args.length
  ? args.map((file, index) => ({
      week: path.basename(file).match(/W\d+/i)?.[0]?.toUpperCase() ?? `W${27 + index}`,
      file,
    }))
  : defaultFiles;
const dataFile = "public/data/initial.json";
const postalDataFile = "public/data/postal-records.json";
const weekStarts = {
  W27: "2026-06-29T00:00:00.000Z",
  W28: "2026-07-06T00:00:00.000Z",
  W29: "2026-07-13T00:00:00.000Z",
  W30: "2026-07-20T00:00:00.000Z",
  W31: "2026-07-27T00:00:00.000Z",
};
const regionNames = {
  NE: "东北区",
  GL: "大湖区",
  FL: "佛州区",
  WE: "美西大区",
  MS: "中南大区",
  TX: "德州大区",
};

const text = (value) => String(value ?? "").trim();
const number = (value) => {
  const parsed = Number(value ?? 0);
  return Number.isFinite(parsed) ? parsed : 0;
};
const postalCode = (value) => {
  const valueText = text(value);
  return /^\d+$/.test(valueText) ? valueText.padStart(5, "0") : valueText;
};

function rowsFrom(file) {
  const workbook = XLSX.readFile(file, { cellDates: true, raw: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  // 大区、站点、DSP、路区在源表中使用纵向合并单元格；把左上角值展开到
  // 合并区域内的每一行，才能得到完整的“路区 × 邮编”原子明细。
  for (const merge of sheet["!merges"] ?? []) {
    const source = sheet[XLSX.utils.encode_cell(merge.s)];
    if (!source) continue;
    for (let row = merge.s.r; row <= merge.e.r; row += 1) {
      for (let column = merge.s.c; column <= merge.e.c; column += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: column });
        if (!sheet[address]) sheet[address] = { ...source };
      }
    }
  }
  return XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
}

function deriveHours(row, attempted) {
  const sourcePph = number(row["PPH"]);
  const perHundredDeliveryMinutes = number(row["百单派件时长(min)"]);
  const perHundredSortMinutes = number(row["百单分拣时长(min)"]);
  const averageDeliveryHours = number(row["司机派件时长(h)"]);
  const averageSortHours = number(row["司机分拣时长(h)"]);
  const averageTransitHours = number(row["司机行驶时长(h)"]);

  const deliveryHours =
    perHundredDeliveryMinutes > 0
      ? (perHundredDeliveryMinutes * attempted) / 6000
      : sourcePph > 0
        ? attempted / sourcePph
        : 0;
  const shiftScale =
    deliveryHours > 0 && averageDeliveryHours > 0
      ? deliveryHours / averageDeliveryHours
      : 0;
  const sortHours =
    shiftScale > 0
      ? averageSortHours * shiftScale
      : perHundredSortMinutes > 0
        ? (perHundredSortMinutes * attempted) / 6000
        : 0;
  const transitHours =
    shiftScale > 0 ? averageTransitHours * shiftScale : 0;

  return {
    sortHours,
    transitHours,
    deliveryHours,
    totalHours: sortHours + transitHours + deliveryHours,
  };
}

const initialData = JSON.parse(fs.readFileSync(dataFile, "utf8"));
const oldPropertyMap = new Map(
  (initialData.properties ?? []).map((item) => [item.route, item]),
);
const oldPostalExact = new Map(
  (initialData.postalProperties ?? []).map((item) => [
    [item.postalCode, item.site, item.dsp, item.route].join("¦"),
    item,
  ]),
);
const oldPostalRoute = new Map(
  (initialData.postalProperties ?? []).map((item) => [
    [item.postalCode, item.site, item.route].join("¦"),
    item,
  ]),
);

const records = [];
const postalRecords = [];
const atomicRows = [];
const fileStats = [];

for (const [weekIndex, entry] of files.entries()) {
  const rows = rowsFrom(entry.file);
  let acceptedRows = 0;
  let skippedTotalRows = 0;
  let skippedDimensionRows = 0;
  for (const row of rows) {
    const regionCode = text(row["大区编码"]).toUpperCase();
    if (regionCode === "总和") {
      skippedTotalRows += 1;
      continue;
    }
    const region = regionNames[regionCode];
    const site = text(row["站点名称"]);
    const dsp = text(row["DSP名称"]);
    const route = text(row["路区名称"]);
    const zip = postalCode(row["邮编"]);
    if (!region || !site || !dsp || !route || !zip) {
      skippedDimensionRows += 1;
      continue;
    }
    const delivered = number(row["妥投量"]);
    const attempted = delivered;
    const hours = deriveHours(row, attempted);
    const base = {
      week: entry.week,
      weekStart: weekStarts[entry.week] ?? "",
      dsp,
      site,
      delivered,
      attempted,
      ...hours,
    };
    records.push({ ...base, route, region });
    postalRecords.push({
      ...base,
      postalCode: zip,
      route,
      region: regionCode,
    });
    atomicRows.push({
      weekIndex,
      week: entry.week,
      regionCode,
      region,
      site,
      dsp,
      route,
      postalCode: zip,
      attempted,
      firstMile: number(row["首单里程 (mi)"]),
      expertPph: number(row["熟手PPH"]),
      populationDensity: number(row["人口密度(人/mi²）"]),
      deliveryExceptionRate: number(row["派送异常率"]),
      dnrRate: number(row["DNR率"]),
    });
    acceptedRows += 1;
  }
  fileStats.push({
    week: entry.week,
    file: entry.file,
    sourceRows: rows.length,
    acceptedRows,
    skippedTotalRows,
    skippedDimensionRows,
  });
}

function buildLatestBuckets(keyOf) {
  const buckets = new Map();
  for (const row of atomicRows) {
    const key = keyOf(row);
    const previous = buckets.get(key);
    if (!previous || row.weekIndex > previous.weekIndex) {
      buckets.set(key, { weekIndex: row.weekIndex, rows: [row] });
    } else if (row.weekIndex === previous.weekIndex) {
      previous.rows.push(row);
    }
  }
  return buckets;
}

const weighted = (rows, field) => {
  const valid = rows.filter((row) => number(row[field]) > 0);
  const weight = valid.reduce((total, row) => total + Math.max(0, row.attempted), 0);
  if (!valid.length) return 0;
  if (weight <= 0) {
    return valid.reduce((total, row) => total + number(row[field]), 0) / valid.length;
  }
  return (
    valid.reduce(
      (total, row) => total + number(row[field]) * Math.max(0, row.attempted),
      0,
    ) / weight
  );
};

const routeBuckets = buildLatestBuckets((row) => row.route);
const properties = [...routeBuckets.entries()]
  .map(([route, bucket]) => {
    const rows = bucket.rows;
    const old = oldPropertyMap.get(route) ?? {};
    const first = rows[0];
    const unified = {
      route,
      businessMode: "",
      sortCode: "",
      transferSite: first.site,
      fleet: first.dsp,
      status: "",
      postalCodes: [...new Set(rows.map((row) => row.postalCode))].join(","),
      addressMix: "",
      safety: "",
      landArea: 0,
      populationDensity: weighted(rows, "populationDensity"),
      isNew: "",
      difficulty: "",
      firstMile: weighted(rows, "firstMile"),
      expertPph: weighted(rows, "expertPph"),
      deliveryExceptionRate: weighted(rows, "deliveryExceptionRate"),
      dnrRate: weighted(rows, "dnrRate"),
      routeUnitPrice: 0,
      routeHourlyWage: 0,
      amazonHourlyMedian: 0,
      salaryCity: "",
    };
    return Object.fromEntries(
      Object.entries(unified).map(([key, value]) => [
        key,
        value === "" || value === 0 ? old[key] ?? value : value,
      ]),
    );
  })
  .sort((a, b) => a.route.localeCompare(b.route));

const postalBuckets = buildLatestBuckets((row) =>
  [row.postalCode, row.site, row.dsp, row.route].join("¦"),
);
const postalProperties = [...postalBuckets.entries()]
  .map(([key, bucket]) => {
    const rows = bucket.rows;
    const first = rows[0];
    const old =
      oldPostalExact.get(key) ??
      oldPostalRoute.get(
        [first.postalCode, first.site, first.route].join("¦"),
      ) ??
      {};
    const routeProperty = oldPropertyMap.get(first.route) ?? {};
    return {
      postalCode: first.postalCode,
      route: first.route,
      site: first.site,
      dsp: first.dsp,
      businessMode: old.businessMode || routeProperty.businessMode || "",
      sortCode: old.sortCode || routeProperty.sortCode || "",
      status: old.status || routeProperty.status || "",
      isNew: old.isNew || routeProperty.isNew || "",
      difficulty: old.difficulty || "",
      firstMile: weighted(rows, "firstMile") || old.firstMile || 0,
      expertPph: weighted(rows, "expertPph") || old.expertPph || 0,
      deliveryExceptionRate:
        weighted(rows, "deliveryExceptionRate") ||
        old.deliveryExceptionRate ||
        0,
      dnrRate: weighted(rows, "dnrRate") || old.dnrRate || 0,
      safety: old.safety || routeProperty.safety || "",
      source: "统一周数据",
    };
  })
  .sort(
    (a, b) =>
      a.route.localeCompare(b.route) ||
      a.postalCode.localeCompare(b.postalCode),
  );

initialData.records = records;
initialData.properties = properties;
initialData.postalProperties = postalProperties;
initialData.meta = {
  ...initialData.meta,
  sourceRows: records.length,
  aggregatedRows: records.length,
  propertyRows: properties.length,
  postalRows: postalRecords.length,
  postalPropertyRows: postalProperties.length,
  generatedAt: new Date().toISOString(),
};

fs.writeFileSync(dataFile, JSON.stringify(initialData), "utf8");
fs.writeFileSync(
  postalDataFile,
  JSON.stringify({ meta: { postalRows: postalRecords.length }, postalRecords }),
  "utf8",
);

console.log(
  JSON.stringify(
    {
      files: fileStats,
      records: records.length,
      postalRecords: postalRecords.length,
      routeProperties: properties.length,
      postalProperties: postalProperties.length,
      postalCostsPreserved: initialData.postalCosts?.length ?? 0,
      weeks: [...new Set(records.map((row) => row.week))],
      regions: [...new Set(postalRecords.map((row) => row.region))].sort(),
    },
    null,
    2,
  ),
);

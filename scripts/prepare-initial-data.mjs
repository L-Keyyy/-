import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx";

const [performanceFile, propertiesFile, difficultyFile] = process.argv.slice(2);

if (!performanceFile || !propertiesFile) {
  console.error(
    "Usage: node scripts/prepare-initial-data.mjs <performance.csv|xlsx> <properties.xlsx> [difficulty.xlsx]",
  );
  process.exit(1);
}

function rowsFromFile(file) {
  const ext = path.extname(file).toLowerCase();
  const workbook =
    ext === ".csv"
      ? XLSX.read(fs.readFileSync(file, "utf8"), {
          type: "string",
          cellDates: true,
        })
      : XLSX.readFile(file, { cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  expandWorksheetRange(sheet);
  return XLSX.utils.sheet_to_json(sheet, {
    defval: null,
  });
}

function expandWorksheetRange(sheet) {
  const addresses = Object.keys(sheet).filter((key) => !key.startsWith("!"));
  if (!addresses.length) return;
  const range = addresses.reduce(
    (current, address) => {
      const cell = XLSX.utils.decode_cell(address);
      current.s.r = Math.min(current.s.r, cell.r);
      current.s.c = Math.min(current.s.c, cell.c);
      current.e.r = Math.max(current.e.r, cell.r);
      current.e.c = Math.max(current.e.c, cell.c);
      return current;
    },
    {
      s: { r: Number.POSITIVE_INFINITY, c: Number.POSITIVE_INFINITY },
      e: { r: 0, c: 0 },
    },
  );
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

const raw = rowsFromFile(performanceFile);
const properties = rowsFromFile(propertiesFile);
const difficultyRows = difficultyFile ? rowsFromFile(difficultyFile) : [];
const grouped = new Map();

for (const row of raw) {
  const route = String(row["路区"] ?? row["路区名称"] ?? "").trim();
  if (!route) continue;
  const key = [
    row["周数"],
    route,
    row["DSP"],
    row["站点"],
    row["大区"],
  ].join("¦");
  const current = grouped.get(key) ?? {
    week: String(row["周数"] ?? ""),
    weekStart:
      row["本周开始日期"] instanceof Date
        ? row["本周开始日期"].toISOString()
        : String(row["本周开始日期"] ?? row["日期"] ?? ""),
    route,
    dsp: String(row["DSP"] ?? ""),
    site: String(row["站点"] ?? ""),
    region: String(row["大区"] ?? ""),
    delivered: 0,
    attempted: 0,
    sortHours: 0,
    transitHours: 0,
    deliveryHours: 0,
    totalHours: 0,
  };

  current.delivered += Number(row["配送量"] ?? 0) || 0;
  current.attempted += Number(row["配送量（加派送失败）"] ?? 0) || 0;
  current.sortHours += Number(row["分拣耗时"] ?? 0) || 0;
  current.transitHours += Number(row["在途耗时"] ?? 0) || 0;
  current.deliveryHours += Number(row["配送耗时"] ?? 0) || 0;
  current.totalHours += Number(row["总时长"] ?? 0) || 0;
  grouped.set(key, current);
}

function normalizeProperty(row) {
  return {
    route: String(row["路区名称"] ?? row["路区"] ?? "").trim(),
    businessMode: String(row["业务模式"] ?? ""),
    sortCode: String(row["分拣码"] ?? ""),
    transferSite: String(row["转运站点"] ?? ""),
    fleet: String(row["车队名称"] ?? ""),
    status: String(row["状态"] ?? ""),
    postalCodes: String(row["邮编"] ?? ""),
    addressMix: String(row["收件地址类型占比"] ?? ""),
    safety: String(row["安全度"] ?? ""),
    landArea: Number(row["陆地面积（mi²）"] ?? 0) || 0,
    populationDensity: Number(row["人口密度（人/mi²）"] ?? 0) || 0,
    isNew: String(row["是否新开"] ?? ""),
    difficulty: String(row["路区难易度"] ?? ""),
    firstMile: Number(row["首单里程（mi）"] ?? 0) || 0,
    expertPph: Number(row["熟手PPH（件）"] ?? 0) || 0,
    deliveryExceptionRate: Number(row["派送异常率"] ?? 0) || 0,
    dnrRate: Number(row["DNR率"] ?? 0) || 0,
    routeUnitPrice: Number(row["路区单均票价"] ?? 0) || 0,
    routeHourlyWage: Number(row["路区时薪"] ?? 0) || 0,
    amazonHourlyMedian:
      Number(
        row["Amazon Flex"] ?? row["亚马逊时薪"] ?? row["Amazon时薪"] ?? 0,
      ) || 0,
    salaryCity: String(row["调研城市名称"] ?? ""),
  };
}

const propertyMap = new Map();
for (const row of [...properties, ...difficultyRows]) {
  const normalized = normalizeProperty(row);
  if (!normalized.route) continue;
  const previous = propertyMap.get(normalized.route);
  if (!previous) {
    propertyMap.set(normalized.route, normalized);
    continue;
  }
  propertyMap.set(
    normalized.route,
    Object.fromEntries(
      Object.entries(normalized).map(([key, value]) => [
        key,
        value === "" || value === 0 ? previous[key] ?? value : value,
      ]),
    ),
  );
}
const normalizedProperties = [...propertyMap.values()];

const output = {
  meta: {
    sourceRows: raw.length,
    aggregatedRows: grouped.size,
    propertyRows: normalizedProperties.length,
    generatedAt: new Date().toISOString(),
  },
  records: [...grouped.values()],
  properties: normalizedProperties,
};

fs.mkdirSync("public/data", { recursive: true });
fs.writeFileSync(
  "public/data/initial.json",
  JSON.stringify(output),
  "utf8",
);
console.log(
  `Prepared ${output.records.length} weekly route rows and ${normalizedProperties.length} property rows.`,
);

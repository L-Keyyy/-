import fs from "node:fs";
import XLSX from "xlsx";

const [
  difficultyFile,
  routePostalFile,
  costFile,
  dataFile = "public/data/initial.json",
  postalDataFile = "public/data/postal-records.json",
] = process.argv.slice(2);

if (!difficultyFile || !routePostalFile || !costFile) {
  console.error(
    "Usage: node scripts/merge-postal-details.mjs <邮编难易度.xlsx> <站点路区邮编对应表.xlsx> <DSP成本.xlsx> [initial.json] [postal-records.json]",
  );
  process.exit(1);
}

const text = (value) => String(value ?? "").trim();
const number = (value) => {
  const parsed = Number(value ?? 0);
  return Number.isFinite(parsed) ? parsed : 0;
};
const rate = (value) => {
  const parsed = number(value);
  return parsed > 1 ? parsed / 100 : parsed;
};
const postalCode = (value) => {
  const valueText = text(value);
  return /^\d+$/.test(valueText) ? valueText.padStart(5, "0") : valueText;
};
const rowsFrom = (file, sheetName) => {
  const workbook = XLSX.readFile(file, { cellDates: true, raw: true });
  const name = sheetName ?? workbook.SheetNames[0];
  if (!workbook.Sheets[name]) {
    throw new Error(`${file} 中未找到工作表：${name}`);
  }
  return XLSX.utils.sheet_to_json(workbook.Sheets[name], {
    defval: null,
    raw: true,
  });
};

const difficultyRows = rowsFrom(difficultyFile);
const routePostalRows = rowsFrom(routePostalFile);
const costRows = rowsFrom(costFile, "1-DSP成本-终端202606");

const postalProperties = new Map();
for (const row of difficultyRows) {
  const zip = postalCode(row["邮编"]);
  const route = text(row["路区名称"]);
  const site = text(row["转运站点"]);
  if (!zip) continue;
  const key = [zip, site, route].join("¦");
  postalProperties.set(key, {
    postalCode: zip,
    route,
    site,
    dsp: text(row["车队名称"]),
    businessMode: text(row["业务模式"]),
    sortCode: text(row["分拣码"]),
    status: text(row["状态"]),
    isNew: text(row["是否新邮编"]),
    difficulty: text(row["邮编难易度"]),
    firstMile: number(row["所属路区首单里程（mi）"]),
    expertPph: number(row["熟手PPH（件）"]),
    deliveryExceptionRate: rate(row["派送异常率"]),
    dnrRate: rate(row["DNR率"]),
    safety: text(row["安全度"]),
    source: "难易度文件",
  });
}

for (const row of routePostalRows) {
  const route = text(row.LUQU);
  const site = text(row.ZD);
  const zips = text(row.YB)
    .split(",")
    .map(postalCode)
    .filter(Boolean);
  for (const zip of zips) {
    const key = [zip, site, route].join("¦");
    if (postalProperties.has(key)) continue;
    postalProperties.set(key, {
      postalCode: zip,
      route,
      site,
      dsp: "",
      businessMode: "",
      sortCode: "",
      status: "",
      isNew: "",
      difficulty: "",
      firstMile: 0,
      expertPph: 0,
      deliveryExceptionRate: 0,
      dnrRate: 0,
      safety: "",
      source: "路区邮编对应表",
    });
  }
}

const postalCosts = new Map();
for (const row of costRows) {
  const zip = postalCode(row["收件邮编"]);
  const route = text(row["路区名称"]);
  const site = text(row["统一目的中心"]);
  const region = text(row["目的大区"]);
  if (!zip || !region) continue;
  const key = [region, site, zip, route].join("¦");
  const current = postalCosts.get(key) ?? {
    postalCode: zip,
    route,
    site,
    region,
    shipmentVolume: 0,
    bookedCost: 0,
    averageDspCost: 0,
  };
  current.shipmentVolume += number(row[" 运单量 "]);
  current.bookedCost += number(row[" 记账成本(未税)(求和) "]);
  postalCosts.set(key, current);
}

for (const cost of postalCosts.values()) {
  cost.averageDspCost =
    cost.shipmentVolume > 0 ? cost.bookedCost / cost.shipmentVolume : 0;
}

const initialData = JSON.parse(fs.readFileSync(dataFile, "utf8"));
initialData.postalProperties = [...postalProperties.values()];
initialData.postalCosts = [...postalCosts.values()];
initialData.meta = {
  ...initialData.meta,
  postalPropertyRows: initialData.postalProperties.length,
  postalCostRows: initialData.postalCosts.length,
  generatedAt: new Date().toISOString(),
};
fs.writeFileSync(dataFile, JSON.stringify(initialData), "utf8");

const performanceZips = new Set(
  JSON.parse(fs.readFileSync(postalDataFile, "utf8")).postalRecords.map(
    (row) => row.postalCode,
  ),
);
const propertyZips = new Set(
  initialData.postalProperties.map((row) => row.postalCode),
);
const costZips = new Set(initialData.postalCosts.map((row) => row.postalCode));

console.log(
  JSON.stringify(
    {
      difficultySourceRows: difficultyRows.length,
      routePostalSourceRows: routePostalRows.length,
      costSourceRows: costRows.length,
      postalPropertyRows: initialData.postalProperties.length,
      postalCostRows: initialData.postalCosts.length,
      performancePostalCodes: performanceZips.size,
      propertyCoverage: [...performanceZips].filter((zip) =>
        propertyZips.has(zip),
      ).length,
      costCoverage: [...performanceZips].filter((zip) => costZips.has(zip))
        .length,
    },
    null,
    2,
  ),
);

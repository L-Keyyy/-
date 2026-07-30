import fs from "node:fs";
import path from "node:path";
import XLSX from "xlsx";

const [performanceFile, propertiesFile] = process.argv.slice(2);

if (!performanceFile || !propertiesFile) {
  console.error(
    "Usage: node scripts/prepare-initial-data.mjs <performance.csv|xlsx> <properties.xlsx>",
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
  return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {
    defval: null,
  });
}

const raw = rowsFromFile(performanceFile);
const properties = rowsFromFile(propertiesFile);
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

const normalizedProperties = properties
  .map((row) => ({
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
  }))
  .filter((row) => row.route);

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

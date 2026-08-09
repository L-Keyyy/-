import fs from "node:fs";
import XLSX from "xlsx";

const [
  postalFile,
  dataFile = "public/data/initial.json",
  postalDataFile = "public/data/postal-records.json",
] = process.argv.slice(2);

if (!postalFile) {
  console.error(
    "Usage: node scripts/merge-postal-data.mjs <postal.csv|xlsx> [initial.json] [postal-records.json]",
  );
  process.exit(1);
}

const isCsv = postalFile.toLowerCase().endsWith(".csv");
const workbook = isCsv
  ? XLSX.read(fs.readFileSync(postalFile, "utf8"), {
      type: "string",
      cellDates: true,
      raw: true,
    })
  : XLSX.readFile(postalFile, { cellDates: true });
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
const grouped = new Map();

const postalCode = (value) => {
  const text = String(value ?? "").trim();
  return /^\d+$/.test(text) ? text.padStart(5, "0") : text;
};

for (const row of rows) {
  const week = String(row["周数"] ?? row["周次"] ?? "").trim();
  const zip = postalCode(row["邮编"] ?? row["收件邮编"]);
  const dsp = String(row["DSP"] ?? row["车队名称"] ?? "").trim();
  const site = String(row["站点"] ?? row["转运站点"] ?? "").trim();
  const region = String(row["大区"] ?? "").trim();
  if (!week || !zip || !region) continue;
  const key = [week, zip, dsp, site, region].join("¦");
  const current = grouped.get(key) ?? {
    week,
    weekStart:
      row["周开始日期"] instanceof Date
        ? row["周开始日期"].toISOString()
        : String(row["周开始日期"] ?? row["本周开始日期"] ?? ""),
    postalCode: zip,
    dsp,
    site,
    region,
    delivered: 0,
    attempted: 0,
    sortHours: 0,
    transitHours: 0,
    deliveryHours: 0,
    totalHours: 0,
  };
  current.delivered += Number(row["配送量"] ?? row["成功配送量"] ?? 0) || 0;
  current.attempted +=
    Number(
      row["配送量（含派送失败）"] ??
        row["配送量（加派送失败）"] ??
        row["配送量"] ??
        0,
    ) || 0;
  current.sortHours += Number(row["分拣耗时"] ?? 0) || 0;
  current.transitHours += Number(row["在途耗时"] ?? 0) || 0;
  current.deliveryHours += Number(row["配送耗时"] ?? 0) || 0;
  current.totalHours += Number(row["总时长"] ?? 0) || 0;
  grouped.set(key, current);
}

const initialData = JSON.parse(fs.readFileSync(dataFile, "utf8"));
const postalRecords = [...grouped.values()];
initialData.meta = {
  ...initialData.meta,
  postalRows: postalRecords.length,
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
      sourceRows: rows.length,
      postalRows: postalRecords.length,
      weeks: [...new Set(postalRecords.map((row) => row.week))],
      postalCodes: new Set(
        postalRecords.map((row) => row.postalCode),
      ).size,
    },
    null,
    2,
  ),
);

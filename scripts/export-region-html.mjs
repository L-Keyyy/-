import fs from "node:fs";

const regionCode = String(process.argv[2] ?? "WE").toUpperCase();
const regionMap = {
  NE: { name: "东北区", source: "东北区" },
  GL: { name: "大湖区", source: "大湖区" },
  FL: { name: "佛州区", source: "佛州区" },
  WE: { name: "美西大区", source: "美西大区" },
  MS: { name: "中南大区", source: "中南大区" },
  TX: { name: "德州大区", source: "德州大区" },
};
const region = regionMap[regionCode];
if (!region) throw new Error(`未知大区：${regionCode}`);

const dataFile = process.argv[3] ?? "public/data/initial.json";
const htmlFile = process.argv[4] ?? "public/PPH周报系统-交互版.html";
const outputFile =
  process.argv[5] ?? `/Users/gl001426/Downloads/${region.name}PPH周报.html`;
const data = JSON.parse(fs.readFileSync(dataFile, "utf8"));
const postalDataFile = process.argv[6] ?? "public/data/postal-records.json";
data.postalRecords = JSON.parse(
  fs.readFileSync(postalDataFile, "utf8"),
).postalRecords;
const records = data.records.filter((row) => row.region === region.source);
const routeKeys = new Set(records.map((row) => row.route));
const postalRecords = (data.postalRecords ?? []).filter(
  (row) => row.region === regionCode,
);
const postalRelationKeys = new Set(
  postalRecords.map((row) =>
    [row.postalCode, row.site, row.dsp, row.route ?? ""].join("¦"),
  ),
);
const postalCodes = new Set(postalRecords.map((row) => row.postalCode));
const scoped = {
  meta: {
    ...data.meta,
    sourceRows: records.length,
    aggregatedRows: records.length,
    propertyRows: data.properties.filter((row) => routeKeys.has(row.route))
      .length,
    postalRows: postalRecords.length,
    postalPropertyRows: (data.postalProperties ?? []).filter((row) =>
      postalRelationKeys.has(
        [row.postalCode, row.site, row.dsp, row.route].join("¦"),
      ),
    ).length,
    postalCostRows: (data.postalCosts ?? []).filter(
      (row) => row.region === regionCode && postalCodes.has(row.postalCode),
    ).length,
    generatedAt: new Date().toISOString(),
  },
  records,
  properties: data.properties.filter((row) => routeKeys.has(row.route)),
  postalRecords,
  postalProperties: (data.postalProperties ?? []).filter((row) =>
    postalRelationKeys.has(
      [row.postalCode, row.site, row.dsp, row.route].join("¦"),
    ),
  ),
  postalCosts: (data.postalCosts ?? []).filter(
    (row) => row.region === regionCode && postalCodes.has(row.postalCode),
  ),
};

let html = fs.readFileSync(htmlFile, "utf8");
const serialized = JSON.stringify(scoped)
  .replace(/&/g, "\\u0026")
  .replace(/</g, "\\u003c")
  .replace(/>/g, "\\u003e")
  .replace(/\u2028/g, "\\u2028")
  .replace(/\u2029/g, "\\u2029");
const script = `<script id="pph-initial-data">window.__PPH_LOCKED_REGION__=${JSON.stringify(regionCode)};window.__PPH_INITIAL_DATA__=${serialized};document.title=${JSON.stringify(`${region.name}PPH周报`)};</script>`;
const marker = '<script id="pph-initial-data">';
const start = html.lastIndexOf(marker);
const end = start >= 0 ? html.indexOf("</script>", start) : -1;
html =
  start >= 0 && end >= 0
    ? `${html.slice(0, start)}${script}${html.slice(end + "</script>".length)}`
    : html.replace(/<\/head>/i, `${script}\n</head>`);
fs.writeFileSync(outputFile, html, "utf8");
console.log(
  JSON.stringify(
    {
      outputFile,
      bytes: fs.statSync(outputFile).size,
      records: scoped.records.length,
      postalRecords: scoped.postalRecords.length,
      properties: scoped.properties.length,
      postalProperties: scoped.postalProperties.length,
    },
    null,
    2,
  ),
);

import fs from "node:fs";
import XLSX from "xlsx";

const [costFile, dataFile = "public/data/initial.json"] = process.argv.slice(2);
if (!costFile) {
  console.error(
    "Usage: node scripts/merge-postal-weight-costs.mjs <DSP成本.xlsx> [initial.json]",
  );
  process.exit(1);
}

const text = (value) => String(value ?? "").trim();
const number = (value) => {
  const parsed = Number(value ?? 0);
  return Number.isFinite(parsed) ? parsed : 0;
};
const postalCode = (value) => {
  const valueText = text(value);
  return /^\d+$/.test(valueText) ? valueText.padStart(5, "0") : valueText;
};
const workbook = XLSX.readFile(costFile, { cellDates: true, raw: true });
const sheetName = workbook.Sheets["1-DSP成本-终端202606"]
  ? "1-DSP成本-终端202606"
  : workbook.SheetNames[0];
const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
  defval: null,
  raw: true,
});
const initialData = JSON.parse(fs.readFileSync(dataFile, "utf8"));
const activeRoutes = new Set(initialData.records.map((row) => row.route));
const postalGroups = new Map();
const weightGroups = new Map();

for (const row of rows) {
  const zip = postalCode(row["收件邮编"]);
  const route = text(row["路区名称"]);
  const site = text(row["统一目的中心"]);
  const region = text(row["目的大区"]);
  const weightBand = text(row["重量段"]);
  const priceType = text(row["价格类型"]);
  if (!zip || !route || !site || !region || !activeRoutes.has(route)) continue;
  const shipmentVolume = number(row[" 运单量 "]);
  const bookedCost = number(row[" 记账成本(未税)(求和) "]);
  const postalKey = [region, site, zip, route].join("¦");
  const postal = postalGroups.get(postalKey) ?? {
    postalCode: zip,
    route,
    site,
    region,
    shipmentVolume: 0,
    bookedCost: 0,
    averageDspCost: 0,
  };
  postal.shipmentVolume += shipmentVolume;
  postal.bookedCost += bookedCost;
  postalGroups.set(postalKey, postal);

  const weightKey = [
    region,
    site,
    zip,
    route,
    weightBand,
    priceType,
  ].join("¦");
  const weight = weightGroups.get(weightKey) ?? {
    postalCode: zip,
    route,
    site,
    region,
    weightBand,
    priceType,
    shipmentVolume: 0,
    bookedCost: 0,
    averageDspCost: 0,
  };
  weight.shipmentVolume += shipmentVolume;
  weight.bookedCost += bookedCost;
  weightGroups.set(weightKey, weight);
}

for (const row of postalGroups.values()) {
  row.averageDspCost =
    row.shipmentVolume > 0 ? row.bookedCost / row.shipmentVolume : 0;
}
for (const row of weightGroups.values()) {
  row.averageDspCost =
    row.shipmentVolume > 0 ? row.bookedCost / row.shipmentVolume : 0;
}

initialData.postalCosts = [...postalGroups.values()];
initialData.postalWeightCosts = [...weightGroups.values()];
initialData.meta = {
  ...initialData.meta,
  postalCostRows: initialData.postalCosts.length,
  postalWeightCostRows: initialData.postalWeightCosts.length,
  generatedAt: new Date().toISOString(),
};
fs.writeFileSync(dataFile, JSON.stringify(initialData), "utf8");

console.log(
  JSON.stringify(
    {
      sourceRows: rows.length,
      activeRoutes: activeRoutes.size,
      postalCostRows: initialData.postalCosts.length,
      postalWeightCostRows: initialData.postalWeightCosts.length,
      coveredRoutes: new Set(initialData.postalCosts.map((row) => row.route))
        .size,
      weightBands: [...new Set(initialData.postalWeightCosts.map((row) => row.weightBand))],
      priceTypes: [...new Set(initialData.postalWeightCosts.map((row) => row.priceType))],
    },
    null,
    2,
  ),
);

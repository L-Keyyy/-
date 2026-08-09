import fs from "node:fs";
import XLSX from "xlsx";

const [salaryFile, dataFile = "public/data/initial.json"] =
  process.argv.slice(2);

if (!salaryFile) {
  console.error(
    "Usage: node scripts/merge-salary-data.mjs <salary.xlsx> [initial.json]",
  );
  process.exit(1);
}

const workbook = XLSX.readFile(salaryFile, { cellDates: true });
const sheet = workbook.Sheets["薪资"] ?? workbook.Sheets[workbook.SheetNames[0]];
const salaryRows = XLSX.utils.sheet_to_json(sheet, { defval: null });
const initialData = JSON.parse(fs.readFileSync(dataFile, "utf8"));
const activeRoutes = new Set(initialData.records.map((row) => row.route));

const median = (values) => {
  const sorted = values
    .map(Number)
    .filter((value) => Number.isFinite(value) && value > 0)
    .sort((a, b) => a - b);
  if (!sorted.length) return 0;
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2
    ? sorted[middle]
    : (sorted[middle - 1] + sorted[middle]) / 2;
};

const salaryGroups = new Map();
for (const row of salaryRows) {
  const route = String(row["路区名称"] ?? row["路区"] ?? "").trim();
  if (!route || !activeRoutes.has(route)) continue;
  const current = salaryGroups.get(route) ?? {
    routeUnitPrices: [],
    routeHourlyWages: [],
    amazonHourlyValues: [],
    cities: new Set(),
  };
  current.routeUnitPrices.push(row["路区单均票价"]);
  current.routeHourlyWages.push(row["路区时薪"]);
  current.amazonHourlyValues.push(
    row["Amazon Flex"] ?? row["亚马逊时薪"] ?? row["Amazon时薪"],
  );
  const city = String(row["调研城市名称"] ?? "").trim();
  if (city) current.cities.add(city);
  salaryGroups.set(route, current);
}

const emptyProperty = (route) => ({
  route,
  businessMode: "",
  sortCode: "",
  transferSite: "",
  fleet: "",
  status: "",
  postalCodes: "",
  addressMix: "",
  safety: "",
  landArea: 0,
  populationDensity: 0,
  isNew: "",
  difficulty: "",
  firstMile: 0,
  expertPph: 0,
  deliveryExceptionRate: 0,
  dnrRate: 0,
  routeUnitPrice: 0,
  routeHourlyWage: 0,
  amazonHourlyMedian: 0,
  salaryCity: "",
});

const propertyMap = new Map(
  initialData.properties.map((item) => [
    item.route,
    {
      ...emptyProperty(item.route),
      ...item,
    },
  ]),
);

for (const [route, group] of salaryGroups) {
  propertyMap.set(route, {
    ...(propertyMap.get(route) ?? emptyProperty(route)),
    routeUnitPrice: median(group.routeUnitPrices),
    routeHourlyWage: median(group.routeHourlyWages),
    amazonHourlyMedian: median(group.amazonHourlyValues),
    salaryCity: [...group.cities].join("; "),
  });
}

initialData.properties = [...propertyMap.values()];
initialData.meta = {
  ...initialData.meta,
  propertyRows: initialData.properties.length,
  generatedAt: new Date().toISOString(),
};

fs.writeFileSync(dataFile, JSON.stringify(initialData), "utf8");
console.log(
  JSON.stringify(
    {
      salaryRows: salaryRows.length,
      matchedRoutes: salaryGroups.size,
      propertyRows: initialData.properties.length,
      amazonMedianRoutes: [...salaryGroups.values()].filter(
        (group) => median(group.amazonHourlyValues) > 0,
      ).length,
    },
    null,
    2,
  ),
);

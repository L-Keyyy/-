import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "vite";

const projectRoot = path.dirname(
  path.dirname(fileURLToPath(import.meta.url)),
);
const distDir = path.join(projectRoot, ".standalone-dist");
const outputFile = path.join(
  projectRoot,
  "public",
  "PPH周报系统-交互版.html",
);

await build({
  configFile: path.join(projectRoot, "vite.standalone.config.ts"),
});

let html = fs.readFileSync(path.join(distDir, "index.html"), "utf8");

const initialPayload = JSON.parse(
  fs.readFileSync(
    path.join(projectRoot, "public", "data", "initial.json"),
    "utf8",
  ),
);
const lockedRegionCode = "WE";
const lockedRegionSource = "美西大区";
initialPayload.records = initialPayload.records.filter(
  (row) => row.region === lockedRegionSource,
);
initialPayload.postalRecords = initialPayload.postalRecords.filter(
  (row) => row.region === lockedRegionCode,
);
const lockedRoutes = new Set([
  ...initialPayload.records.map((row) => row.route),
  ...initialPayload.postalRecords.map((row) => row.route).filter(Boolean),
]);
initialPayload.properties = initialPayload.properties.filter((row) =>
  lockedRoutes.has(row.route),
);
initialPayload.postalProperties = initialPayload.postalProperties.filter(
  (row) => lockedRoutes.has(row.route),
);
initialPayload.postalCosts = initialPayload.postalCosts.filter(
  (row) => row.region === lockedRegionCode,
);
initialPayload.meta = {
  ...initialPayload.meta,
  sourceRows: initialPayload.records.length,
  aggregatedRows: initialPayload.records.length,
  propertyRows: initialPayload.properties.length,
  postalRows: initialPayload.postalRecords.length,
  postalPropertyRows: initialPayload.postalProperties.length,
  postalCostRows: initialPayload.postalCosts.length,
};
const initialData = JSON.stringify(initialPayload)
  .replace(/&/g, "\\u0026")
  .replace(/</g, "\\u003c")
  .replace(/>/g, "\\u003e")
  .replace(/\u2028/g, "\\u2028")
  .replace(/\u2029/g, "\\u2029");
const initialDataScript = `<script id="pph-initial-data">window.__PPH_LOCKED_REGION__=${JSON.stringify(lockedRegionCode)};window.__PPH_INITIAL_DATA__=${initialData};</script>`;
html = html.replace("</head>", `${initialDataScript}\n  </head>`);

const stylesheetMatch = html.match(
  /<link rel="stylesheet" crossorigin href="([^"]+)">/,
);
if (stylesheetMatch) {
  const stylesheetPath = path.join(
    distDir,
    stylesheetMatch[1].replace(/^\.\//, ""),
  );
  const css = fs.readFileSync(stylesheetPath, "utf8");
  html = html.replace(stylesheetMatch[0], () => `<style>${css}</style>`);
}

const scriptMatch = html.match(
  /<script type="module" crossorigin src="([^"]+)"><\/script>/,
);
if (!scriptMatch) {
  throw new Error("Standalone JavaScript bundle was not found.");
}

const scriptPath = path.join(
  distDir,
  scriptMatch[1].replace(/^\.\//, ""),
);
const javascript = fs
  .readFileSync(scriptPath, "utf8")
  .replace(/<\/script>/gi, "<\\\\/script>");
html = html.replace(scriptMatch[0], () => {
  return `<script type="module">${javascript}</script>`;
});

fs.mkdirSync(path.dirname(outputFile), { recursive: true });
fs.writeFileSync(outputFile, html, "utf8");

console.log(
  `Prepared interactive standalone HTML: ${path.relative(projectRoot, outputFile)} (${(
    fs.statSync(outputFile).size /
    1024 /
    1024
  ).toFixed(1)} MB)`,
);

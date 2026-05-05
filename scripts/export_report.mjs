import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

function columnLetterFromIndex(index) {
  let current = index + 1;
  let label = "";
  while (current > 0) {
    const remainder = (current - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    current = Math.floor((current - 1) / 26);
  }
  return label;
}

const [payloadPath, outputPath] = process.argv.slice(2);
if (!payloadPath || !outputPath) {
  throw new Error("Usage: node export_report.mjs <payload.json> <output.xlsx>");
}

const artifactModulePath = process.env.ASSOCIAZIONE_ARTIFACT_TOOL_MODULE;
if (!artifactModulePath) {
  throw new Error("ASSOCIAZIONE_ARTIFACT_TOOL_MODULE non impostata.");
}

const artifactModuleSpecifier = artifactModulePath.startsWith("file://")
  ? artifactModulePath
  : pathToFileURL(artifactModulePath).href;

const { Workbook, SpreadsheetFile } = await import(artifactModuleSpecifier);
const payload = JSON.parse(await fs.readFile(payloadPath, "utf8"));

const workbook = Workbook.create();
const worksheet = workbook.worksheets.add(payload.sheetName || "Report");

const rows = [];
rows.push([payload.title]);
if (payload.subtitle) {
  rows.push([payload.subtitle]);
}
rows.push([`Generato il: ${payload.generatedAt}`]);

for (const item of payload.filters || []) {
  rows.push([`${item.label}: ${item.value}`]);
}

rows.push([]);
rows.push((payload.columns || []).map((column) => column.label));

if ((payload.rows || []).length > 0) {
  rows.push(...payload.rows);
} else {
  rows.push([payload.emptyMessage || "Nessun dato disponibile."]);
}

const maxColumns = Math.max(
  1,
  ...(rows.map((row) => row.length)),
);

const normalizedRows = rows.map((row) => {
  const normalized = [...row];
  while (normalized.length < maxColumns) {
    normalized.push(null);
  }
  return normalized;
});

const lastColumn = columnLetterFromIndex(maxColumns - 1);
const lastRow = normalizedRows.length;
const usedRange = worksheet.getRange(`A1:${lastColumn}${lastRow}`);

usedRange.values = normalizedRows;
usedRange.format.autofitColumns();
usedRange.format.autofitRows();

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

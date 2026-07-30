import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2];
const outputPath = process.argv[3];
const input = await FileBlob.load(inputPath);
const wb = await SpreadsheetFile.importXlsx(input);

const token = wb.worksheets.getOrAdd("Design Tokens");
token.getRange("A1:C10").values = [
  ["Design Tokens", "Hex", "Purpose"],
  ["Base", "#FAF7F1", "Warm ivory base"],
  ["Panel", "#FFFFFF", "Cards and editable surfaces"],
  ["Primary", "#496477", "Primary accent and titles"],
  ["Primary Light", "#DAE5EC", "Chapter and support bands"],
  ["Secondary", "#AC755E", "Warnings and secondary accent"],
  ["Secondary Light", "#EDCCBC", "Changed content"],
  ["Neutral", "#E6DDD0", "Metadata and separators"],
  ["Success", "#C4D8C2", "Completed state"],
  ["Border", "#D9D2C7", "Thin internal boundary"],
];
token.getRange("A1:C1").format = { fill: "#496477", font: { bold: true, color: "#FFFFFF", name: "Aptos" }, horizontalAlignment: "center" };
token.getRange("A2:C10").format = { font: { name: "Aptos", size: 10 }, borders: { preset: "all", style: "thin", color: "#D9D2C7" } };
token.getRange("A1:C10").format.columnWidth = 20;
token.getRange("A1:C1").format.rowHeight = 26;
token.showGridLines = false;
token.freezePanes.freezeRows(1);

for (const [index, kind] of [[1, "dashboard"], [2, "workspace"], [8, "script"], [14, "gantt"]]) {
  const sheet = wb.worksheets.getItemAt(index);
  sheet.showGridLines = false;
  sheet.getRange("A1:N1").format = { fill: "#496477", font: { name: "Aptos", size: 16, bold: true, color: "#FFFFFF" }, verticalAlignment: "center" };
  sheet.getRange("A1:N1").format.rowHeight = 30;
  sheet.getRange("A1:N32").format.borders = { preset: "outside", style: "medium", color: "#496477" };
  sheet.freezePanes.unfreeze();
  sheet.freezePanes.freezeRows(kind === "gantt" ? 3 : 4);
}

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);

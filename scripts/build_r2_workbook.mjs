import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load(process.argv[2]);
const wb = await SpreadsheetFile.importXlsx(input);

const lists = wb.worksheets.getOrAdd("_Lists");
lists.getRange("A1:D9").values = [
  ["Paragraph Types", "Paragraph Statuses", "Asset Statuses", "Priorities"],
  ["Narration", "To Ideate", "To Find", "P0"],
  ["On-screen Text", "Drafting", "Found", "P1"],
  ["Quote / Dialogue", "Finalized", "Used", "P2"],
  ["Visual-only", "Recorded", null, null],
  ["Music Segment", "Assets Matched", null, null],
  ["Transition", "Edited", null, null],
  [null, "Completed", null, null],
  [null, null, null, null],
];
lists.getRange("A1:D1").format = { fill: "#496477", font: { bold: true, color: "#FFFFFF", name: "Aptos" } };
lists.getRange("A1:D9").format.columnWidth = 22;
lists.showGridLines = false;

const scriptStore = wb.worksheets.getOrAdd("_ScriptStore");
scriptStore.getRange("A1:J2").values = [["Script Record ID", "Video ID", "Chapter ID", "Paragraph ID", "Record Type", "Version", "Payload", "Created At", "Updated At", "Active"], [null, null, null, null, null, null, null, null, null, null]];
const recycleStore = wb.worksheets.getOrAdd("_RecycleStore");
recycleStore.getRange("A1:J2").values = [["Recycle Record ID", "Video ID", "Chapter ID", "Paragraph ID", "Record Type", "Original Position", "Payload", "Deleted At", "Restored At", "Active"], [null, null, null, null, null, null, null, null, null, null]];
const migrationLog = wb.worksheets.getOrAdd("_MigrationLog");
migrationLog.getRange("A1:H2").values = [["Migration ID", "Phase", "Run At", "Source Workbook", "Target Workbook", "Status", "Row Count Check", "Notes"], ["R2-INITIAL", "R2", new Date(), "AI_YouTube_Content_System_R1.xlsx", "AI_YouTube_Content_System_R2.xlsx", "Completed", "Preserved", "Additive technical schema only"]];
for (const sheet of [scriptStore, recycleStore, migrationLog]) {
  const used = sheet.getUsedRange();
  used.format = { font: { name: "Aptos", size: 10 }, borders: { preset: "all", style: "thin", color: "#D9D2C7" } };
  sheet.getRange("A1:J1").format = { fill: "#496477", font: { bold: true, color: "#FFFFFF", name: "Aptos" } };
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
}

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(process.argv[3]);

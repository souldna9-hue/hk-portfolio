import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "outputs/news-clipping-template";
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("뉴스 클리핑");
sheet.tabColor = "#1F4E78";
sheet.showGridLines = false;

sheet.getRange("A2:E2").merge();
sheet.getRange("A2").values = [["뉴스 클리핑"]];
sheet.getRange("A3:E3").merge();
sheet.getRange("A3").values = [["기사 정보를 입력해 보관하는 양식입니다. 날짜는 YYYY-MM-DD 형식으로 입력하세요."]];

sheet.getRange("A5:E205").values = [
  ["날짜", "기사 제목", "기사 한줄 요약", "기사 발행 날짜", "뉴스 발행 신문사"],
  ...Array.from({ length: 200 }, () => [null, null, null, null, null]),
];

const header = sheet.getRange("A5:E5");
header.format = {
  fill: "#1F4E78",
  font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#FFFFFF" },
};
sheet.getRange("A6:E205").format = {
  font: { name: "Arial", size: 10, color: "#1F2937" },
  verticalAlignment: "center",
  borders: { preset: "insideHorizontal", style: "thin", color: "#E5E7EB" },
};
sheet.getRange("A6:A205").format.numberFormat = "yyyy-mm-dd";
sheet.getRange("D6:D205").format.numberFormat = "yyyy-mm-dd";
sheet.getRange("A6:A205").format.horizontalAlignment = "center";
sheet.getRange("D6:D205").format.horizontalAlignment = "center";
sheet.getRange("B6:B205").format.wrapText = true;
sheet.getRange("C6:C205").format.wrapText = true;

sheet.getRange("A2:E2").format = {
  font: { name: "Arial", size: 16, bold: true, color: "#17365D" },
  verticalAlignment: "center",
};
sheet.getRange("A3:E3").format = {
  font: { name: "Arial", size: 10, italic: true, color: "#6B7280" },
  verticalAlignment: "center",
};

sheet.getRange("A:A").format.columnWidth = 14;
sheet.getRange("B:B").format.columnWidth = 38;
sheet.getRange("C:C").format.columnWidth = 52;
sheet.getRange("D:D").format.columnWidth = 16;
sheet.getRange("E:E").format.columnWidth = 22;
sheet.getRange("2:2").format.rowHeight = 28;
sheet.getRange("3:3").format.rowHeight = 22;
sheet.getRange("5:5").format.rowHeight = 26;
sheet.getRange("6:205").format.rowHeight = 34;

const table = sheet.tables.add("A5:E205", true, "NewsClippingTable");
table.style = "TableStyleMedium2";
table.showBandedColumns = false;
table.showFilterButton = true;
sheet.freezePanes.freezeRows(5);

workbook.recalculate();
const check = await workbook.inspect({ kind: "table", range: "뉴스 클리핑!A2:E10", include: "values,formulas", tableMaxRows: 10, tableMaxCols: 5 });
console.log(check.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 50 }, summary: "final formula error scan" });
console.log(errors.ndjson);
const preview = await workbook.render({ sheetName: "뉴스 클리핑", range: "A1:E15", scale: 2, format: "png" });
await fs.writeFile(`${outputDir}/preview.png`, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/뉴스_클리핑_양식.xlsx`);

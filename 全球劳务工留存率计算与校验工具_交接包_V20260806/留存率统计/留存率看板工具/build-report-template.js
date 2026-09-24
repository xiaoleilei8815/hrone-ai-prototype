"use strict";

const fs = require("fs");
const path = require("path");
const XlsxPopulate = require("./vendor/xlsx-populate-no-encryption.min.js");

const source = process.argv[2];
if (!source) throw new Error("请传入参考报表路径");

const outputXlsx = path.join(__dirname, "vendor", "retention-report-template.xlsx");
const outputJs = path.join(__dirname, "vendor", "retention-report-template.js");

async function main() {
  const workbook = await XlsxPopulate.fromDataAsync(fs.readFileSync(source));

  for (const sheet of workbook.sheets()) {
    const range = sheet.usedRange();
    if (range) range.value(null);
  }

  const sharedStrings = workbook.sharedStrings();
  sharedStrings._node.children = [];
  sharedStrings._stringArray = [];
  sharedStrings._indexMap = {};

  const chartPath = "xl/charts/chart1.xml";
  const chartFile = workbook._zip.file(chartPath);
  if (chartFile) {
    let chartXml = await chartFile.async("string");
    chartXml = chartXml
      .replace(/<c:strCache>[\s\S]*?<\/c:strCache>/g, '<c:strCache><c:ptCount val="0"/></c:strCache>')
      .replace(/<c:numCache>[\s\S]*?<\/c:numCache>/g, '<c:numCache><c:formatCode>0.0%</c:formatCode><c:ptCount val="0"/></c:numCache>');
    workbook._zip.file(chartPath, chartXml);
  }

  const blob = await workbook.outputAsync();
  const bytes = Buffer.from(await blob.arrayBuffer());
  const encoded = bytes.toString("base64");

  fs.writeFileSync(outputXlsx, bytes);
  fs.writeFileSync(
    outputJs,
    `window.RETENTION_REPORT_TEMPLATE_BASE64 = "${encoded}";\n`,
    "utf8"
  );

  console.log(JSON.stringify({
    source,
    outputXlsx,
    outputJs,
    xlsxBytes: bytes.length,
    base64Chars: encoded.length,
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

"use strict";

const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");
const engine = require("../fulfillment-engine.js");
const exporter = require("../report-exporter.js");

const root = path.resolve(__dirname, "../../../..");
const sourceDir = path.join(root, "需求满足率统计", "欧洲4月-6月需求满足率");
const templatePath = path.join(
  root,
  "需求满足率统计",
  "需求满足率分析",
  "供应商履约率看板工具",
  "vendor",
  "performance-report-template.xlsx",
);
const outputPath = path.join(os.tmpdir(), "欧洲供应商履约率报表_网页导出回归.xlsx");

function exactArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function sourceFiles() {
  return fs
    .readdirSync(sourceDir)
    .filter((name) => name.endsWith(".xlsx") && !name.startsWith("~$"))
    .sort()
    .map((name) => {
      const fullPath = path.join(sourceDir, name);
      const data = fs.readFileSync(fullPath);
      const stat = fs.statSync(fullPath);
      return {
        name,
        size: data.length,
        lastModified: stat.mtimeMs,
        data: exactArrayBuffer(data),
      };
    });
}

async function main() {
  const analysis = await engine.analyzeFiles(sourceFiles());
  const template = fs.readFileSync(templatePath);
  const result = await exporter.buildReport(analysis, template);
  const buffer = Buffer.isBuffer(result.blob)
    ? result.blob
    : Buffer.from(await result.blob.arrayBuffer());
  fs.writeFileSync(outputPath, buffer);

  assert.strictEqual(
    result.fileName,
    "欧洲供应商履约率报表_2026年4-6月_最终严谨版.xlsx",
  );
  assert(buffer.length > 100_000, "导出工作簿体积异常");
  console.log(
    JSON.stringify(
      {
        status: "BUILT",
        output: outputPath,
        bytes: buffer.length,
        fileName: result.fileName,
        expected: {
          supplier_monthly_rows: analysis.supplier_monthly.length,
          display_rows: analysis.display_rows.length,
          supplier_atomic_rows: analysis.supplier_atomic.length,
          unit_atomic_rows: analysis.unit_atomic.length,
        },
      },
      null,
      2,
    ),
  );
}

main().catch((error) => {
  console.error(error.stack || error);
  process.exitCode = 1;
});

"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("../vendor/xlsx.full.min.js");
const engine = require("../fulfillment-engine.js");
const exporter = require("../report-exporter.js");

const root = path.resolve(__dirname, "../../../..");
const sourceDir = path.join(root, "需求满足率统计", "欧洲4月-6月需求满足率");
const templatePath = path.join(__dirname, "..", "vendor", "performance-report-template.xlsx");
const region = "英国区";

function exactArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function allSources() {
  return fs
    .readdirSync(sourceDir)
    .filter((name) => name.endsWith(".xlsx") && !name.startsWith("~$"))
    .sort()
    .map((name) => {
      const fullPath = path.join(sourceDir, name);
      const data = fs.readFileSync(fullPath);
      return {
        name,
        size: data.length,
        lastModified: fs.statSync(fullPath).mtimeMs,
        data: exactArrayBuffer(data),
      };
    });
}

function regionColumn(rows, sheetName) {
  if (sheetName === "基本信息") return 0;
  const header = rows[0] || [];
  const index = header.findIndex((value) => String(value || "").trim() === "需求区域");
  assert(index >= 0, `${sheetName}缺少需求区域`);
  return index;
}

function filterWorkbook(source) {
  const workbook = XLSX.read(source.data, {
    type: "array",
    cellDates: true,
    raw: true,
  });
  const output = XLSX.utils.book_new();
  ["基本信息", "需求详情", "发单详情"].forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      raw: true,
      defval: null,
      blankrows: false,
    });
    const headerCount = sheetName === "基本信息" ? 2 : 1;
    const column = regionColumn(rows, sheetName);
    const filtered = rows
      .slice(0, headerCount)
      .concat(
        rows
          .slice(headerCount)
          .filter((row) => String(row[column] || "").trim() === region),
      );
    XLSX.utils.book_append_sheet(output, XLSX.utils.aoa_to_sheet(filtered), sheetName);
  });
  const data = XLSX.write(output, { type: "array", bookType: "xlsx", cellDates: true });
  return {
    name: source.name.replace(/\.xlsx$/i, "_英国.xlsx"),
    size: data.byteLength,
    lastModified: source.lastModified,
    data,
  };
}

function keyed(rows, dimensions) {
  return new Map(
    rows.map((row) => [
      dimensions.map((dimension) => row[dimension] ?? "").join("\u0001"),
      row,
    ]),
  );
}

function compareSubset(label, actualRows, expectedRows, dimensions, metrics) {
  const actual = keyed(actualRows, dimensions);
  const expected = keyed(expectedRows, dimensions);
  assert.strictEqual(actual.size, expected.size, `${label}键数量`);
  let comparisons = 0;
  expected.forEach((expectedRow, key) => {
    assert(actual.has(key), `${label}缺少键 ${key}`);
    const actualRow = actual.get(key);
    metrics.forEach((metric) => {
      assert(
        Math.abs(Number(actualRow[metric] || 0) - Number(expectedRow[metric] || 0)) <=
          1e-7,
        `${label} ${key} ${metric}`,
      );
      comparisons += 1;
    });
  });
  return comparisons;
}

async function main() {
  const sources = allSources();
  const full = await engine.analyzeFiles(sources);
  const filteredSources = sources.map(filterWorkbook);
  if (process.env.REGION_FIXTURE_DIR) {
    fs.mkdirSync(process.env.REGION_FIXTURE_DIR, { recursive: true });
    filteredSources.forEach((source) => {
      fs.writeFileSync(
        path.join(process.env.REGION_FIXTURE_DIR, source.name),
        Buffer.from(source.data),
      );
    });
  }
  const uk = await engine.analyzeFiles(filteredSources);

  assert.deepStrictEqual(uk.meta.regions, [region]);
  assert.strictEqual(uk.meta.scope_prefix, "英国");
  assert.strictEqual(
    uk.meta.export_filename,
    "英国供应商履约率报表_2026年4-6月_最终严谨版.xlsx",
  );
  assert.strictEqual(uk.meta.coverage.complete, true);
  assert.strictEqual(uk.quality.status, "PASS");
  assert(!uk.display_metrics.some((row) => row.row_type.startsWith("europe_")));

  let comparisons = 0;
  comparisons += compareSubset(
    "英国供应商月度",
    uk.supplier_monthly,
    full.supplier_monthly.filter((row) => row.region === region),
    ["region", "supplier_id", "supplier_name", "job", "month"],
    [
      "issued",
      "dispatched",
      "raw_effective",
      "reasonable_responsibility",
      "adjusted_effective",
      "over_issue_relief",
      "raw_rate",
      "adjusted_rate",
    ],
  );
  comparisons += compareSubset(
    "英国仓库原子",
    uk.unit_atomic,
    full.unit_atomic.filter((row) => row.region === region),
    ["region", "warehouse", "group", "date", "job", "shift"],
    [
      "true_demand",
      "issued",
      "dispatched",
      "warehouse_effective",
      "over_issue",
      "under_issue",
    ],
  );
  comparisons += compareSubset(
    "英国供应商责任原子",
    uk.supplier_atomic,
    full.supplier_atomic.filter((row) => row.region === region),
    [
      "region",
      "warehouse",
      "group",
      "date",
      "job",
      "shift",
      "supplier_id",
      "supplier_name",
    ],
    [
      "issued",
      "dispatched",
      "raw_effective",
      "reasonable_responsibility",
      "adjusted_effective",
      "over_issue_relief",
    ],
  );

  const template = new Uint8Array(fs.readFileSync(templatePath));
  const report = await exporter.buildReport(uk, template);
  assert.strictEqual(report.fileName, uk.meta.export_filename);
  const reportBuffer = Buffer.from(await report.blob.arrayBuffer());
  const workbook = XLSX.read(reportBuffer, { type: "buffer", cellFormula: true });
  assert.deepStrictEqual(workbook.SheetNames, [
    "管理摘要",
    "主指标_原始兑现率",
    "对照_超发校正率",
    "供应商月度明细",
    "仓库满足与发单质量",
    "供应商责任底表",
    "仓库单元底表",
    "三表对账",
    "口径与质量审查",
    "源文件清单",
  ]);
  assert.strictEqual(
    workbook.Sheets["管理摘要"].A1.v,
    "英国供应商履约率管理摘要（2026年4-6月）",
  );

  console.log(
    JSON.stringify(
      {
        status: "PASS",
        files: filteredSources.length,
        comparisons,
        regions: uk.meta.regions,
        suppliers: uk.meta.supplier_count,
        export_filename: report.fileName,
        quality_checks: uk.quality.critical_checks,
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

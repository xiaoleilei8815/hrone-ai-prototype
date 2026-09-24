"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const engine = require("../fulfillment-engine.js");

const root = path.resolve(__dirname, "../../../..");
const sourceDir = path.join(root, "需求满足率统计", "欧洲4月-6月需求满足率");
const expectedPath = path.join(
  root,
  "outputs",
  "europe_supplier_performance_q2_2026_final",
  "analysis.json",
);

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

function keyed(rows, dimensions) {
  return new Map(
    rows.map((row) => [
      dimensions.map((dimension) => row[dimension] ?? "").join("\u0001"),
      row,
    ]),
  );
}

function compareRows(label, actualRows, expectedRows, dimensions, metrics, tolerance = 1e-7) {
  const actual = keyed(actualRows, dimensions);
  const expected = keyed(expectedRows, dimensions);
  assert.strictEqual(actual.size, expected.size, `${label}键数量`);
  let comparisons = 0;
  for (const [key, expectedRow] of expected.entries()) {
    assert(actual.has(key), `${label}缺少键 ${key}`);
    const actualRow = actual.get(key);
    for (const metric of metrics) {
      const left = Number(actualRow[metric] ?? 0);
      const right = Number(expectedRow[metric] ?? 0);
      assert(
        Math.abs(left - right) <= tolerance,
        `${label} ${key} ${metric}: ${left} != ${right}`,
      );
      comparisons += 1;
    }
  }
  return comparisons;
}

async function main() {
  const expected = JSON.parse(fs.readFileSync(expectedPath, "utf8"));
  const actual = await engine.analyzeFiles(sourceFiles());

  assert.deepStrictEqual(actual.clean_totals, expected.clean_totals);
  assert.deepStrictEqual(actual.meta.regions, expected.meta.regions);
  assert.deepStrictEqual(actual.meta.jobs, expected.meta.jobs);
  assert.strictEqual(actual.meta.supplier_count, expected.meta.supplier_count);
  assert.strictEqual(actual.quality.status, "PASS");
  assert.strictEqual(actual.meta.coverage.complete, true);

  let comparisons = 0;
  comparisons += compareRows(
    "供应商月度",
    actual.supplier_monthly,
    expected.supplier_monthly,
    ["region", "supplier_id", "supplier_name", "job", "month"],
    [
      "issued",
      "dispatched",
      "raw_effective",
      "reasonable_responsibility",
      "adjusted_effective",
      "over_issue_relief",
      "dispatch_over_order",
      "raw_rate",
      "adjusted_rate",
      "rate_lift",
    ],
  );
  comparisons += compareRows(
    "仓库原子单元",
    actual.unit_atomic,
    expected.unit_atomic,
    ["region", "warehouse", "group", "date", "job", "shift"],
    [
      "raw_demand",
      "supplement",
      "true_demand",
      "issued",
      "dispatched",
      "issue_covered",
      "warehouse_effective",
      "responsibility_factor",
      "over_issue",
      "zero_demand_issue",
      "under_issue",
      "surplus_dispatch",
    ],
  );
  comparisons += compareRows(
    "供应商责任原子单元",
    actual.supplier_atomic,
    expected.supplier_atomic,
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
      "dispatch_over_order",
    ],
  );

  assert.strictEqual(
    actual.quality.month_region_reconciliation_mismatch_count,
    expected.quality.month_region_reconciliation_mismatch_count,
  );
  assert.strictEqual(
    actual.quality.demand_vs_dispatch_mismatch_count,
    expected.quality.demand_vs_dispatch_mismatch_count,
  );
  assert.strictEqual(
    actual.quality.responsibility_closure_mismatch_count,
    expected.quality.responsibility_closure_mismatch_count,
  );
  assert.strictEqual(
    actual.quality.adjusted_below_raw_count,
    expected.quality.adjusted_below_raw_count,
  );

  console.log(
    JSON.stringify(
      {
        status: "PASS",
        comparisons,
        clean_totals: actual.clean_totals,
        regions: actual.meta.regions,
        suppliers: actual.meta.supplier_count,
        quality_checks: actual.quality.critical_checks,
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

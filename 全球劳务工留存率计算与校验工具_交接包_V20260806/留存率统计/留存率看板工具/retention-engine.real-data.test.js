"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("./vendor/xlsx.full.min.js");
const Engine = require("./retention-engine.js");

const root = path.resolve(__dirname, "../..");
const europe = path.join(root, "留存率统计", "欧洲");
const attendanceFiles = [
  "考勤记录 (欧洲2月) copy.xlsx",
  "欧洲3月考勤记录 copy.xlsx",
  "欧洲4月考勤记录 copy.xlsx",
  "考勤记录（欧洲5月） copy.xlsx",
  "考勤记录 (欧洲6月).xlsx",
];

function readWorkbook(filePath) {
  return XLSX.read(fs.readFileSync(filePath), {
    type: "buffer",
    cellDates: true,
    cellFormula: false,
    cellNF: false,
    cellStyles: false,
  });
}

function balance(summary) {
  return summary.retained + summary.notRetained + summary.immature + summary.pending + summary.exempt;
}

function main() {
  const cohortFile = path.join(europe, "用工管理 (1-3月首次派遣).xlsx");
  const latestFile = path.join(europe, "用工管理 (1-6月内最晚派遣).xlsx");
  const cohort = Engine.parseManagementWorkbook(readWorkbook(cohortFile), XLSX, path.basename(cohortFile));
  const latest = Engine.parseManagementWorkbook(readWorkbook(latestFile), XLSX, path.basename(latestFile));
  const store = Engine.createAttendanceStore();
  const cohortIds = new Set(cohort.records.keys());
  for (const fileName of attendanceFiles) {
    const profile = Engine.mergeAttendanceWorkbook(
      readWorkbook(path.join(europe, fileName)),
      XLSX,
      fileName,
      store,
      cohortIds
    );
    console.log(`attendance ${fileName}: ${profile.rows} rows, ${profile.dateMin}..${profile.dateMax}`);
  }
  Engine.finalizeAttendanceStore(store);

  const baseOptions = {
    cohortStart: Engine.isoToDay("2026-01-01"),
    cohortEnd: Engine.isoToDay("2026-03-31"),
    overrides: new Map(),
  };
  const results90 = Engine.calculate(cohort.records, store, latest.records, {
    ...baseOptions,
    periodDays: 90,
    windowDays: 14,
  });
  const summary90 = Engine.aggregate(results90);
  console.log("90/14", JSON.stringify({
    cohort: summary90.cohort,
    denominator: summary90.denominator,
    retained: summary90.retained,
    notRetained: summary90.notRetained,
    immature: summary90.immature,
    pending: summary90.pending,
    manualReview: summary90.manualReview,
    blindEliminationRisk: summary90.blindEliminationRisk,
    retentionRate: summary90.retentionRate,
  }, null, 2));
  console.log("reasons", summary90.reasons);

  assert.strictEqual(summary90.cohort, 4955, "Q1 cohort count changed");
  assert.strictEqual(summary90.denominator, 4955, "90-day mature denominator changed");
  assert.strictEqual(summary90.retained, 1286, "90-day retained count changed");
  assert.strictEqual(summary90.notRetained, 3669, "90-day not-retained count changed");
  assert.strictEqual(summary90.immature, 0, "Q1 90-day cohort should be mature through June 30");
  assert.strictEqual(summary90.pending, 0, "base data should not have missing L");
  assert.strictEqual(summary90.manualReview, 399, "manual review count changed");
  assert.strictEqual(summary90.blindEliminationRisk, 293, "blind-elimination risk count changed");
  assert.strictEqual(balance(summary90), summary90.cohort, "90/14 outcome balance failed");
  assert.strictEqual(summary90.countries.reduce((sum, row) => sum + row.cohort, 0), summary90.cohort);
  assert.strictEqual(summary90.suppliers.reduce((sum, row) => sum + row.cohort, 0), summary90.cohort);

  const results90T = Engine.calculate(cohort.records, store, latest.records, {
    ...baseOptions,
    periodDays: 90,
    windowDays: 1,
  });
  const summary90T = Engine.aggregate(results90T);
  assert.strictEqual(balance(summary90T), summary90T.cohort, "90/T-only outcome balance failed");
  assert.ok(summary90T.retained <= summary90.retained, "T-only evidence must not retain more people than a 14-day window");

  store.coverageMonths.delete("2026-05");
  const resultsMissingMonth = Engine.calculate(cohort.records, store, latest.records, {
    ...baseOptions,
    periodDays: 90,
    windowDays: 14,
  });
  const summaryMissingMonth = Engine.aggregate(resultsMissingMonth);
  assert.ok(summaryMissingMonth.immature > 0, "a missing attendance month must trigger the maturity gate");
  assert.ok(
    resultsMissingMonth.some((row) => row.missingCoverageMonths.includes("2026-05")),
    "missing month must be preserved in person-level evidence"
  );
  store.coverageMonths.add("2026-05");

  const results365 = Engine.calculate(cohort.records, store, latest.records, {
    ...baseOptions,
    periodDays: 365,
    windowDays: 14,
  });
  const summary365 = Engine.aggregate(results365);
  assert.strictEqual(summary365.immature, summary365.cohort, "365-day cohort must remain right-censored with June data");
  assert.strictEqual(summary365.denominator, 0, "right-censored cohort must not enter the denominator");
  assert.strictEqual(balance(summary365), summary365.cohort, "365/14 outcome balance failed");

  console.log("dynamic checks", JSON.stringify({
    tOnlyRetained: summary90T.retained,
    window14Retained: summary90.retained,
    missingMayImmature: summaryMissingMonth.immature,
    day365Immature: summary365.immature,
  }, null, 2));
  console.log("PASS: real-data engine validation completed");
}

main();

"use strict";

const fs = require("fs");
const path = require("path");
const XLSX = require("../留存率统计/留存率看板工具/vendor/xlsx.full.min.js");
const Engine = require("../留存率统计/留存率看板工具/retention-engine.js");

const workspace = path.resolve(__dirname, "..");
const sourceDir = path.join(workspace, "留存率统计", "美洲亚太");
const cohortFile = path.join(sourceDir, "用工管理（1-3月首次派遣-美洲:亚太）.xlsx");
const latestFile = path.join(sourceDir, "用工管理（1-6月最晚派遣）.xlsx");
const attendanceFiles = [
  "考勤记录（美洲亚太1月）.xlsx",
  "考勤记录 (美洲亚太2月).xlsx",
  "考勤记录（美洲亚太3月）.xlsx",
  "考勤记录（美洲亚太4月）.xlsx",
  "考勤记录 (美洲亚太5月).xlsx",
  "考勤记录（美洲亚太6月）.xlsx",
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

function counts(values) {
  const output = new Map();
  for (const value of values) output.set(value || "（空白）", (output.get(value || "（空白）") || 0) + 1);
  return Array.from(output, ([value, count]) => ({ value, count })).sort((left, right) => right.count - left.count);
}

function main() {
  const cohortWorkbook = readWorkbook(cohortFile);
  const latestWorkbook = readWorkbook(latestFile);
  const cohort = Engine.parseManagementWorkbook(cohortWorkbook, XLSX, path.basename(cohortFile));
  const latest = Engine.parseManagementWorkbook(latestWorkbook, XLSX, path.basename(latestFile));
  const store = Engine.createAttendanceStore();
  const cohortIds = new Set(cohort.records.keys());
  const attendanceProfiles = [];

  for (const fileName of attendanceFiles) {
    attendanceProfiles.push(Engine.mergeAttendanceWorkbook(
      readWorkbook(path.join(sourceDir, fileName)),
      XLSX,
      fileName,
      store,
      cohortIds
    ));
  }
  Engine.finalizeAttendanceStore(store);

  const results = Engine.calculate(cohort.records, store, latest.records, {
    cohortStart: Engine.isoToDay("2026-01-01"),
    cohortEnd: Engine.isoToDay("2026-03-31"),
    periodDays: 90,
    windowDays: 14,
    overrides: new Map(),
  });
  const aggregate = Engine.aggregate(results);
  const records = Array.from(cohort.records.values());
  const rawRegions = counts(records.map((record) => record.region));
  const countries = counts(records.map((record) => record.country));
  const suppliers = counts(records.map((record) => `${record.region} | ${record.supplierName}`));
  const supplierCodes = counts(records.map((record) => `${record.region} | ${record.supplierCode}`));
  const resultRegions = counts(results.map((record) => `${record.region} | ${record.finalResult}`));
  const missingDates = {
    firstDispatch: records.filter((record) => record.firstDispatch == null).length,
    latestDispatch: records.filter((record) => record.latestDispatch == null).length,
    latestEnd: records.filter((record) => record.latestEnd == null).length,
  };

  console.log(JSON.stringify({
    files: {
      cohort: { sheets: cohortWorkbook.SheetNames, profile: cohort.profile },
      latest: { sheets: latestWorkbook.SheetNames, profile: latest.profile },
      attendance: attendanceProfiles,
    },
    cohort: {
      rawRegions,
      mappedCountries: countries,
      supplierCount: new Set(records.map((record) => `${record.region}\u0001${record.supplierName}`)).size,
      suppliers,
      supplierCodes,
      missingDates,
      latestCrossMatches: records.filter((record) => latest.records.has(record.employeeId)).length,
    },
    result: {
      cohort: aggregate.cohort,
      denominator: aggregate.denominator,
      retained: aggregate.retained,
      notRetained: aggregate.notRetained,
      immature: aggregate.immature,
      pending: aggregate.pending,
      manualReview: aggregate.manualReview,
      blindEliminationRisk: aggregate.blindEliminationRisk,
      retentionRate: aggregate.retentionRate,
      regions: resultRegions,
      countryRows: aggregate.countries,
      reasons: aggregate.reasons,
    },
  }, null, 2));
}

main();

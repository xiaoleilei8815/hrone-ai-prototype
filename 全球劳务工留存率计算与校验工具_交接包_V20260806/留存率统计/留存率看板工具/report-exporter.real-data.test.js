"use strict";

const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");
const XLSX = require("./vendor/xlsx.full.min.js");
const Engine = require("./retention-engine.js");
const ReportExporter = require("./report-exporter.js");

const root = path.resolve(__dirname, "../..");
const europe = path.join(root, "留存率统计", "欧洲");
const referencePath = path.join(root, "outputs", "reference_retention_reports", "欧洲劳务工供应商90天留存率报表_2026Q1.xlsx");
const attendanceFiles = [
  "考勤记录 (欧洲1月).xlsx",
  "考勤记录 (欧洲2月) copy.xlsx",
  "欧洲3月考勤记录 copy.xlsx",
  "欧洲4月考勤记录 copy.xlsx",
  "考勤记录（欧洲5月） copy.xlsx",
  "考勤记录 (欧洲6月).xlsx",
];

const expectedSheets = ["统计总览", "国家汇总", "供应商留存率", "人员明细", "异常复核", "数据质量与审查", "口径说明"];
const expectedRefs = {
  "统计总览": "A1:N31",
  "国家汇总": "A1:O8",
  "供应商留存率": "A1:AB55",
  "人员明细": "A1:AP4956",
  "异常复核": "A1:O400",
  "数据质量与审查": "A1:L35",
  "口径说明": "A1:G34",
};
const expectedFormulaCounts = {
  "统计总览": 56,
  "国家汇总": 91,
  "供应商留存率": 1296,
  "人员明细": 39640,
  "异常复核": 0,
  "数据质量与审查": 8,
  "口径说明": 0,
};

function readWorkbook(filePath) {
  return XLSX.read(fs.readFileSync(filePath), {
    type: "buffer",
    cellDates: true,
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
  });
}

function countFormulas(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  let count = 0;
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
      if (cell && cell.f) count += 1;
    }
  }
  return count;
}

function header(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const output = [];
  for (let column = range.s.c; column <= range.e.c; column += 1) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: column })];
    output.push(cell ? cell.v : "");
  }
  return output;
}

function materialDifferenceCount(generated, reference) {
  let differences = 0;
  for (const sheetName of generated.SheetNames) {
    const generatedSheet = generated.Sheets[sheetName];
    const referenceSheet = reference.Sheets[sheetName];
    const range = XLSX.utils.decode_range(generatedSheet["!ref"]);
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let column = range.s.c; column <= range.e.c; column += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: column });
        const generatedCell = generatedSheet[address] || {};
        const referenceCell = referenceSheet[address] || {};
        const generatedValue = generatedCell.v instanceof Date
          ? generatedCell.v.toISOString()
          : generatedCell.v == null || generatedCell.v === "" ? null : generatedCell.v;
        const referenceValue = referenceCell.v instanceof Date
          ? referenceCell.v.toISOString()
          : referenceCell.v == null || referenceCell.v === "" ? null : referenceCell.v;
        const equalNumber = typeof generatedValue === "number" && typeof referenceValue === "number"
          && Math.abs(generatedValue - referenceValue) < 1e-12;
        if ((!equalNumber && generatedValue !== referenceValue) || generatedCell.f !== referenceCell.f) differences += 1;
      }
    }
  }
  return differences;
}

async function main() {
  const cohortFile = path.join(europe, "用工管理 (1-3月首次派遣).xlsx");
  const latestFile = path.join(europe, "用工管理 (1-6月内最晚派遣).xlsx");
  const cohort = Engine.parseManagementWorkbook(readWorkbook(cohortFile), XLSX, path.basename(cohortFile));
  const latest = Engine.parseManagementWorkbook(readWorkbook(latestFile), XLSX, path.basename(latestFile));
  const profiles = { cohort: cohort.profile, latest: latest.profile, attendance: [] };
  const attendanceStore = Engine.createAttendanceStore();
  const cohortIds = new Set(cohort.records.keys());

  for (const fileName of attendanceFiles) {
    profiles.attendance.push(Engine.mergeAttendanceWorkbook(
      readWorkbook(path.join(europe, fileName)),
      XLSX,
      fileName,
      attendanceStore,
      cohortIds
    ));
  }
  Engine.finalizeAttendanceStore(attendanceStore);

  const results = Engine.calculate(cohort.records, attendanceStore, latest.records, {
    cohortStart: Engine.isoToDay("2026-01-01"),
    cohortEnd: Engine.isoToDay("2026-03-31"),
    periodDays: 90,
    windowDays: 14,
    overrides: new Map(),
  });
  const aggregate = Engine.aggregate(results);
  const templatePath = path.join(__dirname, "vendor", "retention-report-template.xlsx");
  const report = await ReportExporter.buildReport({
    aggregate,
    results,
    profiles,
    attendanceStore,
    periodDays: 90,
    windowDays: 14,
    startLabel: "2026-01-01",
    endLabel: "2026-03-31",
    dayToISO: Engine.dayToISO,
  }, fs.readFileSync(templatePath));

  const outputPath = path.join(os.tmpdir(), report.fileName);
  const data = Buffer.isBuffer(report.blob)
    ? report.blob
    : Buffer.from(await report.blob.arrayBuffer());
  fs.writeFileSync(outputPath, data);

  const generated = readWorkbook(outputPath);
  const reference = readWorkbook(referencePath);
  assert.strictEqual(report.fileName, "欧洲劳务工供应商90天留存率报表_2026Q1.xlsx");
  assert.deepStrictEqual(generated.SheetNames, expectedSheets);
  assert.deepStrictEqual(generated.SheetNames, reference.SheetNames);

  for (const sheetName of expectedSheets) {
    assert.strictEqual(generated.Sheets[sheetName]["!ref"], expectedRefs[sheetName], `${sheetName} range changed`);
    assert.strictEqual(generated.Sheets[sheetName]["!ref"], reference.Sheets[sheetName]["!ref"], `${sheetName} range differs from reference`);
    assert.strictEqual(countFormulas(generated.Sheets[sheetName]), expectedFormulaCounts[sheetName], `${sheetName} formula count changed`);
    assert.strictEqual(countFormulas(generated.Sheets[sheetName]), countFormulas(reference.Sheets[sheetName]), `${sheetName} formula count differs from reference`);
  }

  for (const sheetName of ["国家汇总", "供应商留存率", "人员明细", "异常复核"]) {
    assert.deepStrictEqual(header(generated.Sheets[sheetName]), header(reference.Sheets[sheetName]), `${sheetName} header differs from reference`);
  }

  const overview = generated.Sheets["统计总览"];
  assert.strictEqual(overview.A5.v, 4955);
  assert.strictEqual(overview.B5.v, 4955);
  assert.strictEqual(overview.C5.v, 1286);
  assert.strictEqual(overview.D5.v, 3669);
  assert.strictEqual(overview.F5.v, 399);
  assert.strictEqual(overview.G5.v, 293);
  assert.strictEqual(overview.E5.v, reference.Sheets["统计总览"].E5.v);
  assert.strictEqual(materialDifferenceCount(generated, reference), 0, "generated workbook differs from the Q1 reference");

  console.log(JSON.stringify({
    outputPath,
    bytes: data.length,
    sheetRefs: Object.fromEntries(expectedSheets.map((name) => [name, generated.Sheets[name]["!ref"]])),
    formulaCounts: Object.fromEntries(expectedSheets.map((name) => [name, countFormulas(generated.Sheets[name])])),
    summary: { cohort: overview.A5.v, denominator: overview.B5.v, retained: overview.C5.v, notRetained: overview.D5.v },
  }, null, 2));
  console.log("PASS: reference-format report export validation completed");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

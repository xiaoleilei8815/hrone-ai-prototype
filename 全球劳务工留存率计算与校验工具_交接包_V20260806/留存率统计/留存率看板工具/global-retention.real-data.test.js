"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("./vendor/xlsx.full.min.js");
const XlsxPopulate = require("./vendor/xlsx-populate-no-encryption.min.js");
const Engine = require("./retention-engine.js");
const ReportExporter = require("./report-exporter.js");

const root = path.resolve(__dirname, "../..");
const sourceDir = path.join(root, "留存率统计", "美洲亚太");
const outputDir = path.join(root, "outputs", "retention_global_2026Q1");
const cohortName = "用工管理（1-3月首次派遣-美洲:亚太）.xlsx";
const latestName = "用工管理（1-6月最晚派遣）.xlsx";
const attendanceFiles = [
  "考勤记录（美洲亚太1月）.xlsx",
  "考勤记录 (美洲亚太2月).xlsx",
  "考勤记录（美洲亚太3月）.xlsx",
  "考勤记录（美洲亚太4月）.xlsx",
  "考勤记录 (美洲亚太5月).xlsx",
  "考勤记录（美洲亚太6月）.xlsx",
];

function readWorkbook(filePath, formulas) {
  return XLSX.read(fs.readFileSync(filePath), {
    type: "buffer",
    cellDates: true,
    cellFormula: Boolean(formulas),
    cellNF: Boolean(formulas),
    cellStyles: Boolean(formulas),
  });
}

function balance(summary) {
  return summary.retained + summary.notRetained + summary.immature + summary.pending + summary.exempt;
}

function sheetHeader(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const values = [];
  for (let column = range.s.c; column <= range.e.c; column += 1) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: column })];
    values.push(cell ? cell.v : "");
  }
  return values;
}

function cellFormat(sheet, address) {
  return (sheet[address] && sheet[address].z) || "General";
}

function assertCountFormat(sheet, address, label) {
  assert.ok(!cellFormat(sheet, address).includes("%"), `${label} ${address} 不应使用百分比格式`);
}

function assertRateFormat(sheet, address, label) {
  assert.ok(cellFormat(sheet, address).includes("%"), `${label} ${address} 应使用百分比格式`);
}

function mergeAddress(merge) {
  return XLSX.utils.encode_range(merge);
}

function valuesFromChartCache(xml) {
  return Array.from(xml.matchAll(/<c:pt idx="\d+"><c:v>([\s\S]*?)<\/c:v><\/c:pt>/g), (match) => match[1]);
}

async function writeReport(context, scopeLabel, results, template) {
  const aggregate = Engine.aggregate(results);
  const report = await ReportExporter.buildReport({
    ...context,
    scopeLabel,
    results,
    aggregate,
  }, template);
  const data = Buffer.isBuffer(report.blob) ? report.blob : Buffer.from(await report.blob.arrayBuffer());
  const outputPath = path.join(outputDir, report.fileName);
  fs.writeFileSync(outputPath, data);
  const archive = await XlsxPopulate.fromDataAsync(data);
  const chartXml = await archive._zip.file("xl/charts/chart1.xml").async("string");
  return { aggregate, report, outputPath, workbook: readWorkbook(outputPath, true), chartXml, bytes: data.length };
}

async function main() {
  const cohort = Engine.parseManagementWorkbook(
    readWorkbook(path.join(sourceDir, cohortName), false),
    XLSX,
    cohortName
  );
  const latest = Engine.parseManagementWorkbook(
    readWorkbook(path.join(sourceDir, latestName), false),
    XLSX,
    latestName
  );
  const profiles = { cohort: cohort.profile, latest: latest.profile, attendance: [] };
  const attendanceStore = Engine.createAttendanceStore();
  const cohortIds = new Set(cohort.records.keys());
  for (const fileName of attendanceFiles) {
    profiles.attendance.push(Engine.mergeAttendanceWorkbook(
      readWorkbook(path.join(sourceDir, fileName), false),
      XLSX,
      fileName,
      attendanceStore,
      cohortIds
    ));
  }
  Engine.finalizeAttendanceStore(attendanceStore);

  const allResults = Engine.calculate(cohort.records, attendanceStore, latest.records, {
    cohortStart: Engine.isoToDay("2026-01-01"),
    cohortEnd: Engine.isoToDay("2026-03-31"),
    periodDays: 90,
    windowDays: 14,
    overrides: new Map(),
  });
  const all = Engine.aggregate(allResults);
  assert.deepStrictEqual(
    [all.cohort, all.denominator, all.retained, all.notRetained, all.immature, all.pending, all.manualReview, all.blindEliminationRisk],
    [4249, 4249, 911, 3338, 0, 0, 477, 193]
  );
  assert.strictEqual(balance(all), all.cohort);
  assert.strictEqual(all.globalRegions.length, 2);
  assert.strictEqual(all.countryTotals.length, 5);
  assert.strictEqual(all.operatingRegions.length, 14);
  assert.strictEqual(all.suppliers.length, 62);
  assert.strictEqual(allResults.filter((row) => row.globalRegion === "未识别").length, 0);
  assert.strictEqual(all.operatingRegions.reduce((sum, row) => sum + row.cohort, 0), all.cohort);
  assert.strictEqual(all.countryTotals.reduce((sum, row) => sum + row.cohort, 0), all.cohort);
  assert.strictEqual(all.globalRegions.reduce((sum, row) => sum + row.cohort, 0), all.cohort);
  assert.strictEqual(all.suppliers.reduce((sum, row) => sum + row.cohort, 0), all.cohort);

  const americasResults = allResults.filter((row) => row.globalRegion === "美洲");
  const apacResults = allResults.filter((row) => row.globalRegion === "亚太");
  const americas = Engine.aggregate(americasResults);
  const apac = Engine.aggregate(apacResults);
  assert.deepStrictEqual([americas.cohort, americas.retained, americas.notRetained], [4099, 861, 3238]);
  assert.deepStrictEqual([apac.cohort, apac.retained, apac.notRetained], [150, 50, 100]);

  fs.mkdirSync(outputDir, { recursive: true });
  const template = fs.readFileSync(path.join(__dirname, "vendor", "retention-report-template.xlsx"));
  const baseContext = {
    profiles,
    attendanceStore,
    periodDays: 90,
    windowDays: 14,
    startLabel: "2026-01-01",
    endLabel: "2026-03-31",
    dayToISO: Engine.dayToISO,
  };
  const reports = [
    await writeReport(baseContext, "美洲亚太", allResults, template),
    await writeReport(baseContext, "美洲", americasResults, template),
    await writeReport(baseContext, "亚太", apacResults, template),
  ];

  const expectedSheets = ["统计总览", "区域汇总", "供应商留存率", "人员明细", "异常复核", "数据质量与审查", "口径说明"];
  for (const item of reports) {
    assert.deepStrictEqual(item.workbook.SheetNames, expectedSheets);
    assert.deepStrictEqual(sheetHeader(item.workbook.Sheets["区域汇总"]).slice(0, 3), ["大区", "国家", "运营区域"]);
    assert.deepStrictEqual(sheetHeader(item.workbook.Sheets["供应商留存率"]).slice(0, 5), ["大区", "国家", "运营区域", "供应商名称", "供应商编码"]);
    assert.deepStrictEqual(sheetHeader(item.workbook.Sheets["人员明细"]).slice(0, 7), ["大区", "国家", "运营区域", "入职供应商", "入职供应商编码", "姓名", "工号"]);
    assert.deepStrictEqual(sheetHeader(item.workbook.Sheets["异常复核"]).slice(0, 4), ["大区", "国家", "运营区域", "供应商"]);
    const overview = item.workbook.Sheets["统计总览"];
    assert.strictEqual(overview.A5.v, item.aggregate.cohort);
    assert.strictEqual(overview.B5.v, item.aggregate.denominator);
    assert.strictEqual(overview.C5.v, item.aggregate.retained);
    assert.strictEqual(overview.D5.v, item.aggregate.notRetained);
    assert.strictEqual(overview.F5.v, item.aggregate.manualReview);
    assert.strictEqual(overview.G5.v, item.aggregate.blindEliminationRisk);
    assert.ok(/^SUM\(区域汇总!\$D\$2:\$D\$\d+\)$/.test(overview.A5.f));
    assert.strictEqual(item.workbook.Sheets["人员明细"]["!ref"].split(":")[1].startsWith("AQ"), true);

    ["A5", "B5", "C5", "D5", "F5", "G5"].forEach((address) => assertCountFormat(overview, address, "统计总览"));
    assertRateFormat(overview, "E5", "统计总览");

    const areaSheet = item.workbook.Sheets["区域汇总"];
    const areaDataLastRow = XLSX.utils.decode_range(areaSheet["!ref"]).e.r + 1;
    const areas = [];
    for (let row = 2; row <= areaDataLastRow; row += 1) {
      areas.push({
        globalRegion: areaSheet[`A${row}`].v,
        country: areaSheet[`B${row}`].v,
        region: areaSheet[`C${row}`].v,
      });
    }
    const areaKey = (area) => `${area.globalRegion}\u0001${area.country}\u0001${area.region}`;
    assert.deepStrictEqual(
      areas.map(areaKey).sort(),
      item.aggregate.operatingRegions.map(areaKey).sort(),
      "统计总览、区域汇总与聚合结果应覆盖同一组运营区域"
    );
    const areaFirstRow = 9;
    const areaLastRow = areaFirstRow + areas.length - 1;
    areas.forEach((area, index) => {
      const row = areaFirstRow + index;
      assert.strictEqual(overview[`A${row}`].v, area.globalRegion);
      assert.strictEqual(overview[`B${row}`].v, `${area.country}｜${area.region}`);
      ["C", "D", "E"].forEach((column) => assertCountFormat(overview, `${column}${row}`, "统计总览区域表"));
      assertRateFormat(overview, `F${row}`, "统计总览区域表");
      assert.strictEqual(overview[`J${25 + index}`].v, area.region);
      assert.ok(!String(overview[`J${25 + index}`].v).includes("｜"), "图表横轴辅助列不得带国家前缀");
      assertCountFormat(overview, `J${25 + index}`, "图表辅助列");
      assertRateFormat(overview, `K${25 + index}`, "图表辅助列");
    });

    const dataMerges = (overview["!merges"] || []).filter((merge) => merge.s.r <= areaLastRow - 1 && merge.e.r >= areaFirstRow - 1);
    assert.deepStrictEqual(dataMerges.map(mergeAddress), [], "运营区域数据行不得落入合并单元格");
    const coreStart = Math.max(18, 10 + areas.length);
    const merges = new Set((overview["!merges"] || []).map(mergeAddress));
    assert.ok(merges.has(`A${coreStart}:F${coreStart}`), "核心结论标题应只保留一个合并单元格");
    assert.ok(merges.has(`A${coreStart + 1}:F${coreStart + 4}`), "核心结论正文应只保留一个合并区域");
    assert.strictEqual(overview[`A${coreStart}`].v, "核心结论");
    assert.ok(!overview[`B${coreStart}`] || overview[`B${coreStart}`].v == null);
    assert.ok(String(overview[`A${coreStart + 1}`].v).startsWith("本报表按个人T判断"));
    assert.ok(!overview[`B${coreStart + 1}`] || overview[`B${coreStart + 1}`].v == null);

    const categoryBlock = item.chartXml.match(/<c:cat>[\s\S]*?<\/c:cat>/);
    const seriesContainer = item.chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/);
    const seriesBlock = seriesContainer && seriesContainer[0].match(/<c:tx>[\s\S]*?<\/c:tx>/);
    assert.ok(categoryBlock && seriesBlock, "图表 XML 应包含系列与分类缓存");
    assert.ok(item.chartXml.includes('<c:grouping val="clustered"/>'), "柱状图应使用标准 OpenXML 分组枚举");
    assert.ok(!item.chartXml.includes('<c:grouping val="none"/>'), "柱状图不得保留非标准 grouping=none");
    assert.ok(categoryBlock[0].includes(`统计总览!$J$25:$J$${24 + areas.length}`));
    assert.ok(categoryBlock[0].includes(`<c:ptCount val="${areas.length}"/>`));
    assert.deepStrictEqual(valuesFromChartCache(categoryBlock[0]), areas.map((area) => area.region));
    assert.ok(!categoryBlock[0].includes("｜"), "图表横轴分类缓存不得带国家前缀");
    assert.ok(seriesBlock[0].includes('<c:ptCount val="1"/>'));
    assert.deepStrictEqual(valuesFromChartCache(seriesBlock[0]), ["90天留存率"]);

    [2, areaDataLastRow].forEach((row) => {
      ["D", "E", "F", "G", "H", "L", "M", "N", "O", "P"].forEach((column) => assertCountFormat(areaSheet, `${column}${row}`, "区域汇总"));
      ["I", "J", "K"].forEach((column) => assertRateFormat(areaSheet, `${column}${row}`, "区域汇总"));
    });

    const supplierSheet = item.workbook.Sheets["供应商留存率"];
    const supplierLastRow = XLSX.utils.decode_range(supplierSheet["!ref"]).e.r + 1;
    [2, supplierLastRow].forEach((row) => {
      ["F", "G", "H", "I", "J", "N", "O", "P", "Q", "R", "S", "T", "U", "W", "X", "Z", "AA"].forEach((column) => assertCountFormat(supplierSheet, `${column}${row}`, "供应商留存率"));
      ["K", "L", "M", "V", "Y", "AB"].forEach((column) => assertRateFormat(supplierSheet, `${column}${row}`, "供应商留存率"));
    });
  }
  assert.notStrictEqual(reports[0].workbook.Sheets["统计总览"].A18.v, "核心结论", "运营区域列表不应被核心结论覆盖");
  assert.strictEqual(reports[0].workbook.Sheets["统计总览"].A24.v, "核心结论", "合并版核心结论应随区域数量下移");

  assert.strictEqual(reports[0].report.fileName, "美洲亚太劳务工供应商90天留存率报表_2026Q1.xlsx");
  assert.strictEqual(reports[1].report.fileName, "美洲劳务工供应商90天留存率报表_2026Q1.xlsx");
  assert.strictEqual(reports[2].report.fileName, "亚太劳务工供应商90天留存率报表_2026Q1.xlsx");
  console.log(JSON.stringify({
    summary: {
      美洲亚太: [all.cohort, all.retained, all.notRetained, all.retentionRate],
      美洲: [americas.cohort, americas.retained, americas.notRetained, americas.retentionRate],
      亚太: [apac.cohort, apac.retained, apac.notRetained, apac.retentionRate],
    },
    outputs: reports.map((item) => ({ path: item.outputPath, bytes: item.bytes })),
  }, null, 2));
  console.log("PASS: global real-data engine and report validation completed");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

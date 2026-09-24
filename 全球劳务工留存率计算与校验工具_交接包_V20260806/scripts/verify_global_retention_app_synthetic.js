"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const workspace = path.resolve(__dirname, "..");
const toolRoot = path.join(workspace, "留存率统计", "留存率看板工具");
const outputDir = path.join(workspace, "outputs", "global_retention_online_synthetic_verification");
const appUrl = process.argv[2] || "https://ztn.feishuapp.com/app/app_17agg2v5asy/";
const XLSX = require(path.join(toolRoot, "vendor", "xlsx.full.min.js"));
const XlsxPopulate = require(path.join(toolRoot, "vendor", "xlsx-populate-no-encryption.min.js"));

function writeWorkbook(filePath, rows) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  fs.writeFileSync(filePath, XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }));
}

function buildSyntheticFiles() {
  fs.mkdirSync(outputDir, { recursive: true });
  const managementHeader = [
    "工号", "姓名", "状态", "派遣状态", "服务商名称", "服务商编码", "区域",
    "首次派遣日期", "最晚派遣日期", "最近一次派遣结束日期", "派遣结束类型",
    "派遣结束原因", "派遣结束具体原因",
  ];
  const records = [
    ["TEST-EU-001", "测试欧洲", "在职", "派遣中", "Test Europe", "TEST-EU", "英国区", "2026-01-01", "2026-04-10", "", "", "", ""],
    ["TEST-US-001", "测试美国", "在职", "派遣中", "Test USA", "TEST-US", "加州区", "2026-01-01", "2026-04-10", "", "", "", ""],
    ["TEST-CA-001", "测试加拿大", "离职", "已结束", "Test Canada", "TEST-CA", "加拿大区", "2026-01-01", "2026-03-20", "2026-03-20", "", "", ""],
    ["TEST-AU-001", "测试澳洲", "在职", "派遣中", "Test Australia", "TEST-AU", "澳洲区", "2026-01-01", "2026-04-10", "", "", "", ""],
    ["TEST-KR-001", "测试韩国", "在职", "派遣中", "Test Korea", "TEST-KR", "韩国区", "2026-01-01", "2026-04-10", "", "", "", ""],
    ["TEST-JP-001", "测试日本", "在职", "派遣中", "Test Japan", "TEST-JP", "日本区", "2026-01-02", "2026-04-10", "", "", "", ""],
  ];
  const cohortFile = path.join(outputDir, "synthetic-cohort.xlsx");
  const latestFile = path.join(outputDir, "synthetic-latest.xlsx");
  writeWorkbook(cohortFile, [managementHeader, ...records]);
  writeWorkbook(latestFile, [managementHeader, ...records]);

  const attendanceHeader = ["工号", "考勤日期", "时长总计", "确认状态", "考勤类型", "供应商名称", "供应商ID", "区域"];
  const attendanceRows = [];
  const start = Date.UTC(2026, 2, 19);
  const end = Date.UTC(2026, 3, 2);
  for (let time = start; time <= end; time += 86400000) {
    attendanceRows.push(["TEST-EU-001", new Date(time).toISOString().slice(0, 10), 8, "已确认", "派遣考勤", "Test Europe", "TEST-EU", "英国区"]);
  }
  attendanceRows.push(["TEST-US-001", "2026-04-01", 8, "已确认", "派遣考勤", "Test USA", "TEST-US", "加州区"]);
  attendanceRows.push(["TEST-AU-001", "2026-04-01", 8, "已确认", "派遣考勤", "Test Australia", "TEST-AU", "澳洲区"]);
  attendanceRows.push(["TEST-JP-001", "2026-04-02", 8, "已确认", "派遣考勤", "Test Japan", "TEST-JP", "日本区"]);
  attendanceRows.push(["TEST-KR-001", "2026-04-01", 0, "已确认", "派遣考勤", "Test Korea", "TEST-KR", "韩国区"]);
  const attendanceFile = path.join(outputDir, "synthetic-attendance.xlsx");
  writeWorkbook(attendanceFile, [attendanceHeader, ...attendanceRows]);
  return { cohortFile, latestFile, attendanceFile };
}

function readWorkbook(filePath) {
  return XLSX.read(fs.readFileSync(filePath), {
    type: "buffer",
    cellFormula: true,
    cellDates: true,
    cellNF: true,
    cellStyles: true,
  });
}

function cellFormat(sheet, address) {
  return (sheet[address] && sheet[address].z) || "General";
}

function assertGlobalWorkbookLayout(workbook) {
  const overview = workbook.Sheets["统计总览"];
  const areaSheet = workbook.Sheets["区域汇总"];
  const areaLastRow = XLSX.utils.decode_range(areaSheet["!ref"]).e.r + 1;
  const areaCount = areaLastRow - 1;
  const regions = [];
  for (let index = 0; index < areaCount; index += 1) {
    const overviewRow = 9 + index;
    const areaRow = 2 + index;
    const helperRow = 25 + index;
    regions.push(areaSheet[`C${areaRow}`].v);
    ["C", "D", "E"].forEach((column) => {
      assert.ok(!cellFormat(overview, `${column}${overviewRow}`).includes("%"), `${column}${overviewRow} 计数列不得是百分比`);
    });
    assert.ok(cellFormat(overview, `F${overviewRow}`).includes("%"), `F${overviewRow} 留存率必须是百分比`);
    assert.strictEqual(overview[`J${helperRow}`].v, areaSheet[`C${areaRow}`].v);
    assert.ok(!String(overview[`J${helperRow}`].v).includes("｜"));
    assert.ok(cellFormat(overview, `K${helperRow}`).includes("%"));
  }

  const firstDataRow = 8;
  const lastDataRow = 7 + areaCount;
  const dataMerges = (overview["!merges"] || []).filter((merge) => merge.s.r <= lastDataRow && merge.e.r >= firstDataRow);
  assert.deepStrictEqual(dataMerges, [], "运营区域数据行不得被旧合并区域覆盖");
  const coreStart = Math.max(18, 10 + areaCount);
  const merges = new Set((overview["!merges"] || []).map((merge) => XLSX.utils.encode_range(merge)));
  assert.ok(merges.has(`A${coreStart}:F${coreStart}`));
  assert.ok(merges.has(`A${coreStart + 1}:F${coreStart + 4}`));
  assert.strictEqual(overview[`A${coreStart}`].v, "核心结论");
  assert.ok(!overview[`B${coreStart}`] || overview[`B${coreStart}`].v == null);
  assert.ok(!overview[`B${coreStart + 1}`] || overview[`B${coreStart + 1}`].v == null);
  return regions;
}

function chartCacheValues(xml) {
  return Array.from(xml.matchAll(/<c:pt idx="\d+"><c:v>([\s\S]*?)<\/c:v><\/c:pt>/g), (match) => match[1]);
}

async function assertGlobalChart(filePath, expectedRegions) {
  const archive = await XlsxPopulate.fromDataAsync(fs.readFileSync(filePath));
  const xml = await archive._zip.file("xl/charts/chart1.xml").async("string");
  const category = xml.match(/<c:cat>[\s\S]*?<\/c:cat>/);
  assert.ok(category);
  assert.deepStrictEqual(chartCacheValues(category[0]), expectedRegions);
  assert.ok(!category[0].includes("｜"));
  assert.ok(xml.includes('<c:grouping val="clustered"/>'));
  assert.ok(!xml.includes('<c:grouping val="none"/>'));
}

async function waitForCalculation(page) {
  await page.waitForFunction(() => {
    const value = document.querySelector("#progressText");
    return value && (value.textContent.includes("数据就绪") || value.textContent.includes("读取失败"));
  }, { timeout: 180000 });
  const progress = await page.locator("#progressText").textContent();
  if (progress.includes("读取失败")) throw new Error(await page.locator("#systemMessage").textContent());
}

async function readKpis(page) {
  return Promise.all(["kpiCohort", "kpiRetained", "kpiNotRetained", "kpiRate"].map(async (id) => {
    return (await page.locator(`#${id}`).textContent()).trim();
  }));
}

async function exportCurrent(page, expectedName) {
  await page.locator("#exportBtn").click();
  const link = page.locator("#downloadReadyLink");
  await link.waitFor({ state: "visible", timeout: 180000 });
  assert.strictEqual(await link.getAttribute("download"), expectedName);
  const downloadPromise = page.waitForEvent("download", { timeout: 180000 });
  await link.click();
  const download = await downloadPromise;
  const outputPath = path.join(outputDir, download.suggestedFilename());
  await download.saveAs(outputPath);
  return outputPath;
}

async function main() {
  const files = buildSyntheticFiles();
  const consoleErrors = [];
  const pageErrors = [];
  const browser = await chromium.launch({
    headless: true,
    executablePath: process.env.CHROME_PATH || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  });
  try {
    const context = await browser.newContext({ acceptDownloads: true, viewport: { width: 1440, height: 1000 } });
    const page = await context.newPage();
    page.on("console", (message) => {
      if (message.type() === "error") consoleErrors.push(message.text());
    });
    page.on("pageerror", (error) => pageErrors.push(String(error)));
    const response = await page.goto(appUrl, { waitUntil: "networkidle", timeout: 120000 });
    assert(response && response.ok());
    assert.strictEqual(await page.title(), "全球劳务工留存率计算与校验台");
    await page.locator("#cohortFile").setInputFiles(files.cohortFile);
    await page.locator("#latestFile").setInputFiles(files.latestFile);
    await page.locator("#attendanceFiles").setInputFiles(files.attendanceFile);
    await page.locator("#loadBtn").click();
    await waitForCalculation(page);

    assert.deepStrictEqual(await readKpis(page), ["6", "4", "2", "66.7%"]);
    assert.deepStrictEqual(await page.locator("#scopeSelect option").allTextContents(), [
      "全部已上传大区（全球）", "欧洲", "美洲", "亚太",
    ]);
    await page.selectOption("#scopeSelect", "美洲");
    assert.deepStrictEqual(await readKpis(page), ["2", "1", "1", "50.0%"]);
    await page.selectOption("#scopeSelect", "亚太");
    assert.deepStrictEqual(await readKpis(page), ["3", "2", "1", "66.7%"]);
    await page.selectOption("#scopeSelect", "欧洲");
    assert.deepStrictEqual(await readKpis(page), ["1", "1", "0", "100.0%"]);

    await page.selectOption("#scopeSelect", "");
    const chartLabels = await page.locator("#countryBars .bar-label").allTextContents();
    assert.deepStrictEqual(new Set(chartLabels), new Set(["英国区", "加州区", "加拿大区", "澳洲区", "韩国区", "日本区"]));
    assert.ok(chartLabels.every((label) => !label.includes("｜")), "网页图表标签应只显示运营区域");

    await page.evaluate(() => {
      const originalClick = HTMLAnchorElement.prototype.click;
      HTMLAnchorElement.prototype.click = function click() {
        if (this.dataset.autoDownload === "true") return;
        return originalClick.call(this);
      };
    });
    await page.selectOption("#scopeSelect", "欧洲");
    const europePath = await exportCurrent(page, "欧洲劳务工供应商90天留存率报表_20260101-20260102.xlsx");
    assert.deepStrictEqual(readWorkbook(europePath).SheetNames, ["统计总览", "国家汇总", "供应商留存率", "人员明细", "异常复核", "数据质量与审查", "口径说明"]);

    await page.selectOption("#scopeSelect", "亚太");
    const apacPath = await exportCurrent(page, "亚太劳务工供应商90天留存率报表_20260101-20260102.xlsx");
    const apacWorkbook = readWorkbook(apacPath);
    assert.deepStrictEqual(apacWorkbook.SheetNames, ["统计总览", "区域汇总", "供应商留存率", "人员明细", "异常复核", "数据质量与审查", "口径说明"]);
    await assertGlobalChart(apacPath, assertGlobalWorkbookLayout(apacWorkbook));

    await page.selectOption("#scopeSelect", "");
    const outputPath = await exportCurrent(page, "全球劳务工供应商90天留存率报表_20260101-20260102.xlsx");
    const workbook = readWorkbook(outputPath);
    assert.deepStrictEqual(workbook.SheetNames, ["统计总览", "区域汇总", "供应商留存率", "人员明细", "异常复核", "数据质量与审查", "口径说明"]);
    assert.deepStrictEqual([workbook.Sheets["统计总览"].A5.v, workbook.Sheets["统计总览"].C5.v, workbook.Sheets["统计总览"].D5.v], [6, 4, 2]);
    assert.deepStrictEqual([workbook.Sheets["区域汇总"].A1.v, workbook.Sheets["区域汇总"].B1.v, workbook.Sheets["区域汇总"].C1.v], ["大区", "国家", "运营区域"]);
    await assertGlobalChart(outputPath, assertGlobalWorkbookLayout(workbook));
    assert.deepStrictEqual(pageErrors, []);
    assert.deepStrictEqual(consoleErrors, []);
    console.log(JSON.stringify({ appUrl: page.url(), downloadedReports: [europePath, apacPath, outputPath], employees: 6, retained: 4 }, null, 2));
    console.log("PASS: public global app synthetic end-to-end validation completed");
  } finally {
    await browser.close();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const workspace = path.resolve(__dirname, "..");
const toolRoot = path.join(workspace, "留存率统计", "留存率看板工具");
const sourceDir = path.join(workspace, "留存率统计", "美洲亚太");
const outputDir = path.join(workspace, "outputs", "global_retention_app_verification");
const referencePath = path.join(
  workspace,
  "outputs",
  "retention_global_2026Q1",
  "美洲亚太劳务工供应商90天留存率报表_2026Q1.xlsx"
);
const appUrl = process.argv[2] || "https://ztn.feishuapp.com/app/app_17agg2v5asy/";
const XLSX = require(path.join(toolRoot, "vendor", "xlsx.full.min.js"));

const cohortFile = path.join(sourceDir, "用工管理（1-3月首次派遣-美洲:亚太）.xlsx");
const latestFile = path.join(sourceDir, "用工管理（1-6月最晚派遣）.xlsx");
const attendanceFiles = [
  "考勤记录（美洲亚太1月）.xlsx",
  "考勤记录 (美洲亚太2月).xlsx",
  "考勤记录（美洲亚太3月）.xlsx",
  "考勤记录（美洲亚太4月）.xlsx",
  "考勤记录 (美洲亚太5月).xlsx",
  "考勤记录（美洲亚太6月）.xlsx",
].map((name) => path.join(sourceDir, name));

function readWorkbook(filePath) {
  return XLSX.read(fs.readFileSync(filePath), {
    type: "buffer",
    cellDates: true,
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
  });
}

function comparableValue(value) {
  if (value instanceof Date) return value.toISOString();
  return value == null || value === "" ? null : value;
}

function materialDifferenceCount(actual, expected) {
  assert.deepStrictEqual(actual.SheetNames, expected.SheetNames, "工作表名称或顺序不同");
  let differences = 0;
  for (const sheetName of actual.SheetNames) {
    const actualSheet = actual.Sheets[sheetName];
    const expectedSheet = expected.Sheets[sheetName];
    assert.strictEqual(actualSheet["!ref"], expectedSheet["!ref"], `${sheetName}数据范围不同`);
    const range = XLSX.utils.decode_range(actualSheet["!ref"]);
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let column = range.s.c; column <= range.e.c; column += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: column });
        const actualCell = actualSheet[address] || {};
        const expectedCell = expectedSheet[address] || {};
        const actualValue = comparableValue(actualCell.v);
        const expectedValue = comparableValue(expectedCell.v);
        const equalNumber = typeof actualValue === "number"
          && typeof expectedValue === "number"
          && Math.abs(actualValue - expectedValue) < 1e-12;
        if ((!equalNumber && actualValue !== expectedValue) || actualCell.f !== expectedCell.f) differences += 1;
      }
    }
  }
  return differences;
}

function cellFormat(sheet, address) {
  return (sheet[address] && sheet[address].z) || "General";
}

function assertGlobalOverviewLayout(workbook) {
  const overview = workbook.Sheets["统计总览"];
  const areaSheet = workbook.Sheets["区域汇总"];
  const areaLastRow = XLSX.utils.decode_range(areaSheet["!ref"]).e.r + 1;
  const areaCount = areaLastRow - 1;
  const regions = [];
  for (let index = 0; index < areaCount; index += 1) {
    const row = 9 + index;
    const areaRow = 2 + index;
    regions.push(areaSheet[`C${areaRow}`].v);
    ["C", "D", "E"].forEach((column) => assert.ok(!cellFormat(overview, `${column}${row}`).includes("%")));
    assert.ok(cellFormat(overview, `F${row}`).includes("%"));
    assert.strictEqual(overview[`J${25 + index}`].v, areaSheet[`C${areaRow}`].v);
    assert.ok(!String(overview[`J${25 + index}`].v).includes("｜"));
    assert.ok(cellFormat(overview, `K${25 + index}`).includes("%"));
  }
  const coreStart = Math.max(18, 10 + areaCount);
  const merges = new Set((overview["!merges"] || []).map((merge) => XLSX.utils.encode_range(merge)));
  assert.ok(merges.has(`A${coreStart}:F${coreStart}`));
  assert.ok(merges.has(`A${coreStart + 1}:F${coreStart + 4}`));
  assert.strictEqual(overview[`A${coreStart}`].v, "核心结论");
  assert.ok(!overview[`B${coreStart}`] || overview[`B${coreStart}`].v == null);
  return regions;
}

async function waitForCalculation(page) {
  const deadline = Date.now() + 12 * 60 * 1000;
  let previous = "";
  while (Date.now() < deadline) {
    const progress = (await page.locator("#progressText").textContent()) || "";
    if (progress !== previous) {
      console.log(`progress: ${progress}`);
      previous = progress;
    }
    if (progress.includes("数据就绪")) return;
    if (progress.includes("读取失败")) {
      throw new Error((await page.locator("#systemMessage").textContent()) || progress);
    }
    await page.waitForTimeout(2500);
  }
  throw new Error("页面读取美洲亚太OTWS文件超时");
}

async function readKpis(page) {
  return {
    cohort: (await page.locator("#kpiCohort").textContent()).trim(),
    denominator: (await page.locator("#kpiDenominator").textContent()).trim(),
    retained: (await page.locator("#kpiRetained").textContent()).trim(),
    notRetained: (await page.locator("#kpiNotRetained").textContent()).trim(),
    rate: (await page.locator("#kpiRate").textContent()).trim(),
  };
}

async function main() {
  fs.mkdirSync(outputDir, { recursive: true });
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
    assert(response && response.ok(), `页面HTTP状态异常：${response && response.status()}`);
    await page.locator("h1").waitFor({ state: "visible", timeout: 30000 });
    assert.strictEqual(await page.title(), "全球劳务工留存率计算与校验台");
    assert.match(await page.locator("#libraryStatus").textContent(), /已就绪/);

    await page.locator("#cohortFile").setInputFiles(cohortFile);
    await page.locator("#latestFile").setInputFiles(latestFile);
    await page.locator("#attendanceFiles").setInputFiles(attendanceFiles);
    await page.locator("#loadBtn").click();
    await waitForCalculation(page);

    const combined = await readKpis(page);
    assert.deepStrictEqual(combined, {
      cohort: "4,249",
      denominator: "4,249",
      retained: "911",
      notRetained: "3,338",
      rate: "21.4%",
    });
    assert.deepStrictEqual(await page.locator("#scopeSelect option").allTextContents(), [
      "全部已上传大区（美洲亚太）",
      "美洲",
      "亚太",
    ]);
    assert.strictEqual((await page.locator("#countryCount").textContent()).trim(), "14个运营区域");
    const webChartLabels = await page.locator("#countryBars .bar-label").allTextContents();
    assert.ok(webChartLabels.every((label) => !label.includes("｜")), "网页图表应只显示运营区域");

    await page.selectOption("#scopeSelect", "美洲");
    const americas = await readKpis(page);
    assert.deepStrictEqual(americas, {
      cohort: "4,099",
      denominator: "4,099",
      retained: "861",
      notRetained: "3,238",
      rate: "21.0%",
    });
    assert.strictEqual((await page.locator("#countryCount").textContent()).trim(), "11个运营区域");

    await page.selectOption("#scopeSelect", "亚太");
    const apac = await readKpis(page);
    assert.deepStrictEqual(apac, {
      cohort: "150",
      denominator: "150",
      retained: "50",
      notRetained: "100",
      rate: "33.3%",
    });
    assert.strictEqual((await page.locator("#countryCount").textContent()).trim(), "3个运营区域");

    await page.selectOption("#scopeSelect", "");
    assert.deepStrictEqual(await readKpis(page), combined);
    await page.evaluate(() => {
      const originalClick = HTMLAnchorElement.prototype.click;
      HTMLAnchorElement.prototype.click = function click() {
        if (this.dataset.autoDownload === "true") return;
        return originalClick.call(this);
      };
    });
    await page.locator("#exportBtn").click();
    const downloadLink = page.locator("#downloadReadyLink");
    await downloadLink.waitFor({ state: "visible", timeout: 180000 });
    assert.strictEqual(
      await downloadLink.getAttribute("download"),
      "美洲亚太劳务工供应商90天留存率报表_2026Q1.xlsx"
    );
    const downloadPromise = page.waitForEvent("download", { timeout: 180000 });
    await downloadLink.click();
    const download = await downloadPromise;
    const downloadPath = path.join(outputDir, download.suggestedFilename());
    await download.saveAs(downloadPath);
    const downloadedWorkbook = readWorkbook(downloadPath);
    const referenceWorkbook = readWorkbook(referencePath);
    const differences = materialDifferenceCount(downloadedWorkbook, referenceWorkbook);
    assert.strictEqual(differences, 0, "页面导出报表与已验证参考报表存在值或公式差异");
    assert.deepStrictEqual(new Set(webChartLabels), new Set(assertGlobalOverviewLayout(downloadedWorkbook)));

    const desktopScreenshot = path.join(outputDir, "全球留存率工具桌面验收.png");
    await page.screenshot({ path: desktopScreenshot, fullPage: false });
    await page.setViewportSize({ width: 390, height: 844 });
    await page.waitForTimeout(500);
    const bodyOverflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    assert.ok(bodyOverflow <= 1, `移动端页面出现${bodyOverflow}px全局横向溢出`);
    const mobileScreenshot = path.join(outputDir, "全球留存率工具移动端验收.png");
    await page.screenshot({ path: mobileScreenshot, fullPage: false });
    assert.deepStrictEqual(pageErrors, [], `页面脚本异常：${pageErrors.join(" | ")}`);
    assert.deepStrictEqual(consoleErrors, [], `控制台异常：${consoleErrors.join(" | ")}`);

    console.log(JSON.stringify({
      appUrl: page.url(),
      combined,
      americas,
      apac,
      downloadedReport: downloadPath,
      materialDifferences: differences,
      desktopScreenshot,
      mobileScreenshot,
    }, null, 2));
    console.log("PASS: global retention app end-to-end validation completed");
  } finally {
    await browser.close();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

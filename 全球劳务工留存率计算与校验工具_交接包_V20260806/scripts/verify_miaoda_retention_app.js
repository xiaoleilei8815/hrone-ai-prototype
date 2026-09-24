"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const workspace = path.resolve(__dirname, "..");
const toolRoot = path.join(workspace, "留存率统计", "留存率看板工具");
const europe = path.join(workspace, "留存率统计", "欧洲");
const outputDir = path.join(workspace, "outputs", "miaoda_retention_verification");
const appUrl = process.argv[2] || "https://ztn.feishuapp.com/app/app_17agg2v5asy";
const referencePath = path.join(
  workspace,
  "outputs",
  "reference_retention_reports",
  "欧洲劳务工供应商90天留存率报表_2026Q1.xlsx"
);
const XLSX = require(path.join(toolRoot, "vendor", "xlsx.full.min.js"));

const cohortFile = path.join(europe, "用工管理 (1-3月首次派遣).xlsx");
const latestFile = path.join(europe, "用工管理 (1-6月内最晚派遣).xlsx");
const attendanceFiles = [
  "考勤记录 (欧洲1月).xlsx",
  "考勤记录 (欧洲2月) copy.xlsx",
  "欧洲3月考勤记录 copy.xlsx",
  "欧洲4月考勤记录 copy.xlsx",
  "考勤记录（欧洲5月） copy.xlsx",
  "考勤记录 (欧洲6月).xlsx",
].map((name) => path.join(europe, name));

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

function comparableStyle(style) {
  if (!style) return null;
  const keys = Object.keys(style);
  if (keys.length === 1 && style.patternType === "none") return null;
  return style;
}

function workbookDifferenceCounts(actual, expected) {
  let valueOrFormula = 0;
  let semanticCellStyle = 0;
  let blankNoFillSerialization = 0;
  let sheetLayout = 0;

  assert.deepStrictEqual(actual.SheetNames, expected.SheetNames, "工作表名称或顺序不同");
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
        if ((!equalNumber && actualValue !== expectedValue) || actualCell.f !== expectedCell.f) {
          valueOrFormula += 1;
        }
        const rawActualStyle = JSON.stringify(actualCell.s || null);
        const rawExpectedStyle = JSON.stringify(expectedCell.s || null);
        if (rawActualStyle !== rawExpectedStyle) {
          const actualStyle = JSON.stringify(comparableStyle(actualCell.s));
          const expectedStyle = JSON.stringify(comparableStyle(expectedCell.s));
          if (actualStyle === expectedStyle) blankNoFillSerialization += 1;
          else semanticCellStyle += 1;
        }
      }
    }

    const layoutKeys = ["!cols", "!rows", "!merges", "!autofilter", "!margins", "!protect"];
    for (const key of layoutKeys) {
      if (JSON.stringify(actualSheet[key] || null) !== JSON.stringify(expectedSheet[key] || null)) {
        sheetLayout += 1;
      }
    }
  }
  return { valueOrFormula, semanticCellStyle, blankNoFillSerialization, sheetLayout };
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
    await page.waitForTimeout(3000);
  }
  throw new Error("线上页面读取OTWS文件超时");
}

async function main() {
  fs.mkdirSync(outputDir, { recursive: true });
  const consoleErrors = [];
  const pageErrors = [];
  const requestFailures = [];
  const browser = await chromium.launch({
    headless: true,
    executablePath: process.env.CHROME_PATH
      || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  });

  try {
    const context = await browser.newContext({
      acceptDownloads: true,
      viewport: { width: 1440, height: 1000 },
    });
    const page = await context.newPage();
    page.on("console", (message) => {
      if (message.type() === "error") consoleErrors.push(message.text());
    });
    page.on("pageerror", (error) => pageErrors.push(String(error)));
    page.on("requestfailed", (request) => {
      requestFailures.push(`${request.url()} :: ${request.failure() && request.failure().errorText}`);
    });

    const response = await page.goto(appUrl, { waitUntil: "networkidle", timeout: 120000 });
    assert(response && response.ok(), `线上页面HTTP状态异常：${response && response.status()}`);
    await page.locator("h1").waitFor({ state: "visible", timeout: 30000 });
    assert.strictEqual(await page.title(), "全球劳务工留存率计算与校验台");
    assert.strictEqual((await page.locator("h1").textContent()).trim(), "全球劳务工留存率计算与校验台");
    assert.match(await page.locator("#libraryStatus").textContent(), /已就绪/);

    await page.locator("#cohortFile").setInputFiles(cohortFile);
    await page.locator("#latestFile").setInputFiles(latestFile);
    await page.locator("#attendanceFiles").setInputFiles(attendanceFiles);
    await page.locator("#loadBtn").click();
    await waitForCalculation(page);

    const baseline = {
      cohort: (await page.locator("#kpiCohort").textContent()).trim(),
      denominator: (await page.locator("#kpiDenominator").textContent()).trim(),
      retained: (await page.locator("#kpiRetained").textContent()).trim(),
      notRetained: (await page.locator("#kpiNotRetained").textContent()).trim(),
      rate: (await page.locator("#kpiRate").textContent()).trim(),
      review: (await page.locator("#kpiReview").textContent()).trim(),
    };
    assert.deepStrictEqual(baseline, {
      cohort: "4,955",
      denominator: "4,955",
      retained: "1,286",
      notRetained: "3,669",
      rate: "26.0%",
      review: "399",
    });

    await page.locator('#periodSegments button[data-days="custom"]').click();
    await page.locator("#periodCustom").fill("45");
    await page.locator('#windowSegments button[data-days="custom"]').click();
    await page.locator("#windowCustom").fill("10");
    await page.locator("#calculateBtn").click();
    assert.strictEqual((await page.locator("#countryChartTitle").textContent()).trim(), "欧洲成熟队列45天留存率");
    assert.strictEqual((await page.locator("#kpiDenominator").textContent()).trim(), "4,955");
    const customRate = (await page.locator("#kpiRate").textContent()).trim();
    assert.notStrictEqual(customRate, baseline.rate, "自定义45天/10日窗口未改变计算结果");

    await page.locator('#periodSegments button[data-days="90"]').click();
    await page.locator('#windowSegments button[data-days="14"]').click();
    await page.locator("#calculateBtn").click();
    assert.strictEqual((await page.locator("#kpiRetained").textContent()).trim(), "1,286");

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
    assert.match(await page.locator("#systemMessage").textContent(), /点击顶部【下载已生成报表】/);
    assert.match(await downloadLink.getAttribute("href"), /^blob:/);
    assert.strictEqual(
      await downloadLink.getAttribute("download"),
      "欧洲劳务工供应商90天留存率报表_2026Q1.xlsx"
    );

    const downloadPromise = page.waitForEvent("download", { timeout: 180000 });
    await downloadLink.click();
    const download = await downloadPromise;
    const downloadPath = path.join(outputDir, download.suggestedFilename());
    await download.saveAs(downloadPath);
    assert.strictEqual(download.suggestedFilename(), "欧洲劳务工供应商90天留存率报表_2026Q1.xlsx");
    assert(fs.statSync(downloadPath).size > 1000000, "线上导出文件体积异常");

    const actualWorkbook = readWorkbook(downloadPath);
    const expectedWorkbook = readWorkbook(referencePath);
    const differences = workbookDifferenceCounts(actualWorkbook, expectedWorkbook);
    assert.strictEqual(differences.valueOrFormula, 0, "单元格值或公式与参考报表不同");
    assert.strictEqual(differences.semanticCellStyle, 0, "存在可见或有业务意义的单元格样式差异");
    assert.strictEqual(differences.sheetLayout, 0, "工作表行列、合并或页面布局与参考报表不同");

    const screenshotPath = path.join(outputDir, "妙搭线上验收截图.png");
    await page.screenshot({ path: screenshotPath, fullPage: true });
    assert.deepStrictEqual(pageErrors, [], `页面脚本异常：${pageErrors.join(" | ")}`);
    assert.deepStrictEqual(consoleErrors, [], `控制台异常：${consoleErrors.join(" | ")}`);
    assert.deepStrictEqual(requestFailures, [], `静态资源请求失败：${requestFailures.join(" | ")}`);

    console.log(JSON.stringify({
      appUrl: page.url(),
      baseline,
      custom45Day10WindowRate: customRate,
      downloadedReport: downloadPath,
      downloadedBytes: fs.statSync(downloadPath).size,
      workbookDifferences: differences,
      blockedAutoDownloadFallback: "PASS",
      screenshot: screenshotPath,
      publicSessionHadNoLoginRedirect: page.url().includes("ztn.feishuapp.com/app/app_17agg2v5asy"),
    }, null, 2));
    console.log("PASS: 妙搭线上端到端验收完成");
  } finally {
    await browser.close();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

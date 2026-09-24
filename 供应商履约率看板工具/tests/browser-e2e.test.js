"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");
const { launchOptions } = require("./playwright-launch.js");

const workspaceRoot = path.resolve(__dirname, "../../../..");
const sourceDir = path.join(
  workspaceRoot,
  "需求满足率统计",
  "欧洲4月-6月需求满足率",
);
const sourceFiles = fs
  .readdirSync(sourceDir)
  .filter((name) => name.endsWith(".xlsx") && !name.startsWith("~$"))
  .sort()
  .map((name) => path.join(sourceDir, name));
const baseUrl = process.env.APP_URL || "http://127.0.0.1:8765/";
const desktopScreenshot = "/private/tmp/supplier-dashboard-desktop.png";
const mobileScreenshot = "/private/tmp/supplier-dashboard-mobile.png";
const downloadedReport = "/private/tmp/欧洲供应商履约率报表_浏览器回归.xlsx";

async function analyze(page) {
  await page.goto(baseUrl, { waitUntil: "networkidle" });
  await page.setInputFiles("#fileInput", sourceFiles);
  await page.click("#analyzeButton");
  await page.waitForSelector("#statusPanel.is-success", { timeout: 60000 });
  await page.waitForFunction(() => !document.querySelector("#dashboard").hidden);
}

async function canvasPaintedPixels(page, selector) {
  return page.$eval(selector, (canvas) => {
    const context = canvas.getContext("2d");
    const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
    let count = 0;
    for (let index = 3; index < pixels.length; index += 4) {
      if (pixels[index] > 0) count += 1;
    }
    return { count, width: canvas.width, height: canvas.height };
  });
}

async function main() {
  const browser = await chromium.launch(launchOptions());
  const consoleErrors = [];
  const pageErrors = [];
  try {
    const desktopContext = await browser.newContext({
      viewport: { width: 1440, height: 1000 },
      acceptDownloads: true,
      locale: "zh-CN",
    });
    const page = await desktopContext.newPage();
    page.on("console", (message) => {
      if (message.type() === "error") consoleErrors.push(message.text());
    });
    page.on("pageerror", (error) => pageErrors.push(error.message));
    await analyze(page);

    assert.strictEqual(
      await page.textContent("#reportTitle"),
      "欧洲供应商履约率报表（2026年4-6月）",
    );
    assert((await page.textContent("#scopeRegion")).includes("英国区"));
    assert.strictEqual(await page.textContent("#qualityBadge"), "PASS");
    assert((await page.textContent("#kpiGrid")).includes("76.9%"));
    assert((await page.textContent("#kpiGrid")).includes("70.3%"));

    const trendPixels = await canvasPaintedPixels(page, "#trendChart");
    const scatterPixels = await canvasPaintedPixels(page, "#supplierScatter");
    assert(trendPixels.count > 500, "趋势图未形成有效像素");
    assert(scatterPixels.count > 500, "散点图未形成有效像素");

    await page.click('[data-tab="suppliers"]');
    assert((await page.locator("#supplierTableBody tr").count()) > 20);
    await page.click('[data-supplier-metric="adjusted"]');
    assert((await page.textContent("#supplierTableHead")).includes("合理责任"));
    await page.click('[data-tab="warehouse"]');
    assert((await page.locator("#warehouseTableBody tr").count()) > 6);
    await page.click('[data-tab="quality"]');
    assert.strictEqual(await page.locator(".quality-check.fail").count(), 0);
    assert.strictEqual(await page.locator(".quality-check").count(), 6);
    await page.click('[data-tab="overview"]');

    await page.selectOption("#regionFilter", "英国区");
    assert((await page.textContent("#supplierPreviewMeta")).includes("英国"));
    await page.selectOption("#regionFilter", "__all__");

    const downloadPromise = page.waitForEvent("download", { timeout: 60000 });
    await page.click("#exportButton");
    const download = await downloadPromise;
    assert.strictEqual(
      download.suggestedFilename(),
      "欧洲供应商履约率报表_2026年4-6月_最终严谨版.xlsx",
    );
    await download.saveAs(downloadedReport);
    assert(fs.statSync(downloadedReport).size > 100000);

    await page.screenshot({ path: desktopScreenshot });
    await desktopContext.close();

    const mobileContext = await browser.newContext({
      viewport: { width: 390, height: 844 },
      locale: "zh-CN",
    });
    const mobile = await mobileContext.newPage();
    mobile.on("console", (message) => {
      if (message.type() === "error") consoleErrors.push(message.text());
    });
    mobile.on("pageerror", (error) => pageErrors.push(error.message));
    await analyze(mobile);
    const overflow = await mobile.evaluate(
      () => document.documentElement.scrollWidth - window.innerWidth,
    );
    assert(overflow <= 1, `移动端出现${overflow}px全局横向溢出`);
    const mobileTrend = await canvasPaintedPixels(mobile, "#trendChart");
    assert(mobileTrend.count > 300, "移动端趋势图未形成有效像素");
    await mobile.screenshot({ path: mobileScreenshot });
    await mobileContext.close();

    assert.deepStrictEqual(pageErrors, [], `页面脚本错误：${pageErrors.join(" | ")}`);
    assert.deepStrictEqual(consoleErrors, [], `浏览器控制台错误：${consoleErrors.join(" | ")}`);
    console.log(
      JSON.stringify(
        {
          status: "PASS",
          source_files: sourceFiles.length,
          desktop: {
            trend_pixels: trendPixels,
            scatter_pixels: scatterPixels,
            screenshot: desktopScreenshot,
          },
          mobile: {
            trend_pixels: mobileTrend,
            screenshot: mobileScreenshot,
            horizontal_overflow_px: overflow,
          },
          download: downloadedReport,
          page_errors: pageErrors.length,
          console_errors: consoleErrors.length,
        },
        null,
        2,
      ),
    );
  } finally {
    await browser.close();
  }
}

main().catch((error) => {
  console.error(error.stack || error);
  process.exitCode = 1;
});

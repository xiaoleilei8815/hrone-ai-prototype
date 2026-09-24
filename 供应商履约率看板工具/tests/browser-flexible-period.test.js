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
const yearFixtureDir =
  process.env.FLEX_FIXTURE_DIR || "/private/tmp/supplier-dashboard-12m-sources";
const baseUrl = process.env.APP_URL || "http://127.0.0.1:8765/";

function sourceFiles(directory, predicate = () => true) {
  return fs
    .readdirSync(directory)
    .filter((name) => name.endsWith(".xlsx") && !name.startsWith("~$") && predicate(name))
    .sort()
    .map((name) => path.join(directory, name));
}

async function uploadAndAnalyze(page, files, timeout = 120000) {
  await page.goto(baseUrl, { waitUntil: "networkidle" });
  await page.setInputFiles("#fileInput", files);
  await page.click("#analyzeButton");
  await page.waitForSelector("#statusPanel.is-success", { timeout });
  await page.waitForFunction(() => !document.querySelector("#dashboard").hidden);
}

async function downloadReport(page, expectedName, outputPath, timeout = 120000) {
  const downloadPromise = page.waitForEvent("download", { timeout });
  await page.click("#exportButton");
  const download = await downloadPromise;
  assert.strictEqual(download.suggestedFilename(), expectedName);
  await download.saveAs(outputPath);
  assert(fs.statSync(outputPath).size > 100000, "导出文件体积异常");
  assert.strictEqual(await page.locator("#downloadLink").isVisible(), true);
}

async function retryPreparedDownload(page, expectedName) {
  const downloadPromise = page.waitForEvent("download", { timeout: 20000 });
  await page.click("#exportButton");
  const download = await downloadPromise;
  assert.strictEqual(download.suggestedFilename(), expectedName);
}

async function paintedPixels(page, selector) {
  return page.$eval(selector, (canvas) => {
    const pixels = canvas
      .getContext("2d")
      .getImageData(0, 0, canvas.width, canvas.height).data;
    let count = 0;
    for (let index = 3; index < pixels.length; index += 4) {
      if (pixels[index] > 0) count += 1;
    }
    return count;
  });
}

async function main() {
  const aprilFiles = sourceFiles(sourceDir, (name) => /\(4\./.test(name));
  const yearFiles = sourceFiles(yearFixtureDir);
  assert.strictEqual(aprilFiles.length, 2, "4月源文件数量");
  assert.strictEqual(yearFiles.length, 28, "12个月测试源文件数量");

  const browser = await chromium.launch(launchOptions());
  const errors = [];
  try {
    const context = await browser.newContext({
      viewport: { width: 1440, height: 1000 },
      acceptDownloads: true,
      locale: "zh-CN",
    });
    const page = await context.newPage();
    page.on("pageerror", (error) => errors.push(error.message));
    page.on("console", (message) => {
      if (message.type() === "error") errors.push(message.text());
    });

    await uploadAndAnalyze(page, aprilFiles);
    assert.strictEqual(
      await page.textContent("#reportTitle"),
      "欧洲供应商履约率报表（2026年4月）",
    );
    assert.strictEqual(await page.locator("#allocationChart .allocation-row").count(), 1);
    await page.click('[data-tab="suppliers"]');
    assert.strictEqual(await page.locator("#supplierTableHead th").count(), 9);
    assert((await page.textContent("#supplierTableHead")).includes("4月"));
    await downloadReport(
      page,
      "欧洲供应商履约率报表_2026年4月_最终严谨版.xlsx",
      "/private/tmp/欧洲供应商履约率报表_浏览器单月回归.xlsx",
    );
    await retryPreparedDownload(
      page,
      "欧洲供应商履约率报表_2026年4月_最终严谨版.xlsx",
    );

    await uploadAndAnalyze(page, yearFiles, 180000);
    assert.strictEqual(
      await page.textContent("#reportTitle"),
      "欧洲供应商履约率报表（2026年4月1日-2027年3月30日）",
    );
    assert.strictEqual(await page.locator("#allocationChart .allocation-row").count(), 12);
    await page.click('[data-tab="suppliers"]');
    assert.strictEqual(await page.locator("#supplierTableHead th").count(), 20);
    const headerText = await page.textContent("#supplierTableHead");
    assert(headerText.includes("2026年4月"));
    assert(headerText.includes("2027年3月"));
    await page.click('[data-tab="overview"]');
    assert((await paintedPixels(page, "#trendChart")) > 500);
    await downloadReport(
      page,
      "欧洲供应商履约率报表_2026年4月1日-2027年3月30日_最终严谨版.xlsx",
      "/private/tmp/欧洲供应商履约率报表_浏览器12个月回归.xlsx",
      240000,
    );
    await page.screenshot({ path: "/private/tmp/supplier-dashboard-12m.png" });

    assert.deepStrictEqual(errors, [], `页面错误：${errors.join(" | ")}`);
    await context.close();
    console.log(
      JSON.stringify(
        {
          status: "PASS",
          single_month_files: aprilFiles.length,
          twelve_month_files: yearFiles.length,
          twelve_month_columns: 12,
          downloads: 3,
          browser_errors: errors.length,
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

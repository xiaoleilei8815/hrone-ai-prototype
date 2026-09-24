"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("../vendor/xlsx.full.min.js");
const { chromium } = require("playwright");
const { launchOptions } = require("./playwright-launch.js");

const sourceDir =
  process.env.REGION_FIXTURE_DIR || "/private/tmp/supplier-dashboard-uk-sources";
const sourceFiles = fs
  .readdirSync(sourceDir)
  .filter((name) => name.endsWith(".xlsx"))
  .sort()
  .map((name) => path.join(sourceDir, name));
const baseUrl = process.env.APP_URL || "http://127.0.0.1:8765/";
const downloadPath = "/private/tmp/英国供应商履约率报表_线上区域回归.xlsx";

async function main() {
  assert.strictEqual(sourceFiles.length, 7, "英国区域源文件数量");
  const browser = await chromium.launch(launchOptions());
  const errors = [];
  try {
    const context = await browser.newContext({
      viewport: { width: 1280, height: 900 },
      acceptDownloads: true,
      locale: "zh-CN",
    });
    const page = await context.newPage();
    page.on("pageerror", (error) => errors.push(error.message));
    page.on("console", (message) => {
      if (message.type() === "error") errors.push(message.text());
    });
    await page.goto(baseUrl, { waitUntil: "networkidle" });
    await page.setInputFiles("#fileInput", sourceFiles);
    await page.click("#analyzeButton");
    await page.waitForSelector("#statusPanel.is-success", { timeout: 60000 });

    assert.strictEqual(
      await page.textContent("#reportTitle"),
      "英国供应商履约率报表（2026年4-6月）",
    );
    assert.strictEqual(await page.textContent("#scopeRegion"), "英国区");
    assert.strictEqual(await page.textContent("#scopeSupplier"), "7 家");
    assert.strictEqual(await page.textContent("#qualityBadge"), "PASS");
    assert.strictEqual(await page.locator("#regionFilter option").count(), 1);
    assert.strictEqual(
      await page.locator("#regionFilter option").textContent(),
      "英国区",
    );
    assert.strictEqual(await page.locator("#supplierPreviewBody tr").count(), 7);

    const downloadPromise = page.waitForEvent("download", { timeout: 60000 });
    await page.click("#exportButton");
    const download = await downloadPromise;
    assert.strictEqual(
      download.suggestedFilename(),
      "英国供应商履约率报表_2026年4-6月_最终严谨版.xlsx",
    );
    await download.saveAs(downloadPath);
    const workbook = XLSX.read(fs.readFileSync(downloadPath), {
      type: "buffer",
      cellFormula: true,
    });
    assert.strictEqual(
      workbook.Sheets["管理摘要"].A1.v,
      "英国供应商履约率管理摘要（2026年4-6月）",
    );
    assert.deepStrictEqual(errors, []);
    await context.close();
    console.log(
      JSON.stringify(
        {
          status: "PASS",
          source_files: sourceFiles.length,
          report_title: "英国供应商履约率报表（2026年4-6月）",
          supplier_count: 7,
          quality: "PASS",
          download: downloadPath,
          errors: errors.length,
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

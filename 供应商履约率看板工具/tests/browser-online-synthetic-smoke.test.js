"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("../vendor/xlsx.full.min.js");
const { chromium } = require("playwright");
const { launchOptions } = require("./playwright-launch.js");

const appUrl = process.env.APP_URL;
const fixtureDir =
  process.env.SYNTHETIC_FIXTURE_DIR || "/private/tmp/supplier-dashboard-global-synthetic";
const outputPath = "/private/tmp/全球供应商履约率报表_线上虚构样本验收.xlsx";

async function main() {
  assert(appUrl, "缺少 APP_URL");
  const sourceFiles = fs
    .readdirSync(fixtureDir)
    .filter((name) => name.endsWith(".xlsx"))
    .sort()
    .map((name) => path.join(fixtureDir, name));
  assert.strictEqual(sourceFiles.length, 3);

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
    await page.goto(appUrl, { waitUntil: "networkidle", timeout: 60000 });
    await page.setInputFiles("#fileInput", sourceFiles);
    await page.click("#analyzeButton");
    await page.waitForSelector("#statusPanel.is-success", { timeout: 60000 });

    assert((await page.textContent("#reportTitle")).startsWith("全球供应商履约率报表"));
    assert.strictEqual(await page.textContent("#scopeMarket"), "欧洲、美洲、亚太");
    assert.strictEqual(await page.textContent("#scopeSupplier"), "4 家");
    assert.strictEqual(await page.textContent("#qualityBadge"), "PASS");
    assert.strictEqual(await page.locator("#marketFilter option").count(), 4);
    assert.strictEqual(await page.locator("#regionFilter option").count(), 5);

    const downloadPromise = page.waitForEvent("download", { timeout: 120000 });
    await page.click("#exportButton");
    const download = await downloadPromise;
    assert(download.suggestedFilename().startsWith("全球供应商履约率报表_"));
    assert(download.suggestedFilename().endsWith("_最终严谨版.xlsx"));
    await download.saveAs(outputPath);
    assert(fs.statSync(outputPath).size > 40000);

    const workbook = XLSX.read(fs.readFileSync(outputPath), {
      type: "buffer",
      cellFormula: true,
    });
    assert(workbook.Sheets["管理摘要"].A1.v.startsWith("全球供应商履约率管理摘要"));
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets["主指标_原始兑现率"], {
      header: 1,
      raw: true,
      defval: null,
    });
    const labels = new Set(rows.slice(4).map((row) => row[0]).filter(Boolean));
    ["欧洲总计", "美洲总计", "亚太总计", "全球总计"].forEach((label) => {
      assert(labels.has(label), `线上导出缺少${label}`);
    });
    assert.deepStrictEqual(errors, [], `线上页面错误：${errors.join(" | ")}`);
    await context.close();
    console.log(
      JSON.stringify(
        {
          status: "PASS",
          app_url: appUrl,
          source_files: sourceFiles.length,
          markets: 3,
          regions: 4,
          suppliers: 4,
          download: outputPath,
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

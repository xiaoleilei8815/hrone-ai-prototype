"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("../vendor/xlsx.full.min.js");
const { chromium } = require("playwright");
const { launchOptions } = require("./playwright-launch.js");

const workspaceRoot = path.resolve(__dirname, "../../../..");
const europeDir = path.join(
  workspaceRoot,
  "需求满足率统计",
  "欧洲4月-6月需求满足率",
);
const americasApacDir = path.join(
  workspaceRoot,
  "需求满足率统计",
  "美洲亚太4月-6月需求满足率",
);
const baseUrl = process.env.APP_URL || "http://127.0.0.1:8765/";

function files(directory) {
  return fs
    .readdirSync(directory)
    .filter((name) => name.endsWith(".xlsx") && !name.startsWith("~$"))
    .sort()
    .map((name) => path.join(directory, name));
}

async function analyze(page, sourceFiles, timeout = 120000) {
  await page.goto(baseUrl, { waitUntil: "networkidle" });
  await page.setInputFiles("#fileInput", sourceFiles);
  await page.click("#analyzeButton");
  await page.waitForSelector("#statusPanel.is-success", { timeout });
  await page.waitForFunction(() => !document.querySelector("#dashboard").hidden);
}

async function download(page, expectedName, outputPath, timeout = 240000) {
  const downloadPromise = page.waitForEvent("download", { timeout });
  await page.click("#exportButton");
  const item = await downloadPromise;
  assert.strictEqual(item.suggestedFilename(), expectedName);
  await item.saveAs(outputPath);
  assert(fs.statSync(outputPath).size > 500000, "全球化报表体积异常");
}

async function paintedPixels(page, selector) {
  return page.$eval(selector, (canvas) => {
    const data = canvas
      .getContext("2d")
      .getImageData(0, 0, canvas.width, canvas.height).data;
    let count = 0;
    for (let index = 3; index < data.length; index += 4) {
      if (data[index] > 0) count += 1;
    }
    return count;
  });
}

function mainSheetLabels(outputPath) {
  const workbook = XLSX.read(fs.readFileSync(outputPath), {
    type: "buffer",
    cellFormula: true,
  });
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets["主指标_原始兑现率"], {
    header: 1,
    raw: true,
    defval: null,
  });
  return {
    managementTitle: workbook.Sheets["管理摘要"].A1.v,
    labels: new Set(rows.slice(4).map((row) => row[0]).filter(Boolean)),
  };
}

async function main() {
  const europeFiles = files(europeDir);
  const americasApacFiles = files(americasApacDir);
  assert.strictEqual(europeFiles.length, 7);
  assert.strictEqual(americasApacFiles.length, 7);

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

    await analyze(page, americasApacFiles);
    assert.strictEqual(
      await page.textContent("#reportTitle"),
      "美洲亚太供应商履约率报表（2026年4-6月）",
    );
    assert.strictEqual(await page.textContent("#scopeMarket"), "美洲、亚太");
    assert.strictEqual(await page.textContent("#scopeSupplier"), "60 家");
    assert.strictEqual(await page.textContent("#qualityBadge"), "PASS");
    assert.strictEqual(await page.locator("#marketFilter option").count(), 3);
    assert.strictEqual(await page.locator("#regionFilter option").count(), 12);
    assert((await page.textContent("#kpiGrid")).includes("38.1%"));
    assert((await page.textContent("#kpiGrid")).includes("53.2%"));
    assert((await paintedPixels(page, "#trendChart")) > 500);
    assert((await paintedPixels(page, "#supplierScatter")) > 500);

    await page.selectOption("#marketFilter", "亚太");
    assert.strictEqual(await page.locator("#regionFilter option").count(), 1);
    assert.strictEqual(await page.locator("#regionFilter option").textContent(), "澳洲区");
    assert((await page.textContent("#kpiGrid")).includes("39.2%"));
    await page.click('[data-tab="suppliers"]');
    assert.strictEqual(await page.locator("#supplierTableBody tr").count(), 6);

    await page.selectOption("#marketFilter", "美洲");
    assert.strictEqual(await page.locator("#regionFilter option").count(), 11);
    assert.strictEqual(await page.locator("#supplierTableBody tr").count(), 54);
    await page.click('[data-tab="overview"]');
    await page.selectOption("#marketFilter", "__all__");

    const combinedOutput = "/private/tmp/美洲亚太供应商履约率报表_浏览器回归.xlsx";
    await download(
      page,
      "美洲亚太供应商履约率报表_2026年4-6月_最终严谨版.xlsx",
      combinedOutput,
    );
    const combinedWorkbook = mainSheetLabels(combinedOutput);
    assert.strictEqual(
      combinedWorkbook.managementTitle,
      "美洲亚太供应商履约率管理摘要（2026年4-6月）",
    );
    ["美洲总计", "亚太总计", "美洲亚太总计"].forEach((label) =>
      assert(combinedWorkbook.labels.has(label), `主表缺少${label}`),
    );
    await page.screenshot({ path: "/private/tmp/supplier-dashboard-americas-apac.png" });

    await analyze(page, europeFiles.concat(americasApacFiles), 180000);
    assert.strictEqual(
      await page.textContent("#reportTitle"),
      "全球供应商履约率报表（2026年4-6月）",
    );
    assert.strictEqual(await page.textContent("#scopeMarket"), "欧洲、美洲、亚太");
    assert.strictEqual(await page.textContent("#scopeSupplier"), "104 家");
    assert.strictEqual(await page.locator("#marketFilter option").count(), 4);
    assert.strictEqual(await page.locator("#regionFilter option").count(), 18);
    await page.selectOption("#marketFilter", "欧洲");
    assert.strictEqual(await page.locator("#regionFilter option").count(), 7);
    await page.selectOption("#marketFilter", "__all__");

    const globalOutput = "/private/tmp/全球供应商履约率报表_浏览器回归.xlsx";
    await download(
      page,
      "全球供应商履约率报表_2026年4-6月_最终严谨版.xlsx",
      globalOutput,
      300000,
    );
    const globalWorkbook = mainSheetLabels(globalOutput);
    assert.strictEqual(
      globalWorkbook.managementTitle,
      "全球供应商履约率管理摘要（2026年4-6月）",
    );
    ["欧洲总计", "美洲总计", "亚太总计", "全球总计"].forEach((label) =>
      assert(globalWorkbook.labels.has(label), `全球主表缺少${label}`),
    );

    assert.deepStrictEqual(errors, [], `页面错误：${errors.join(" | ")}`);
    await context.close();

    const mobileErrors = [];
    const mobileContext = await browser.newContext({
      viewport: { width: 390, height: 844 },
      acceptDownloads: true,
      locale: "zh-CN",
    });
    const mobilePage = await mobileContext.newPage();
    mobilePage.on("pageerror", (error) => mobileErrors.push(error.message));
    mobilePage.on("console", (message) => {
      if (message.type() === "error") mobileErrors.push(message.text());
    });
    await analyze(mobilePage, americasApacFiles);
    const mobileLayout = await mobilePage.evaluate(() => ({
      viewportWidth: window.innerWidth,
      documentWidth: document.documentElement.scrollWidth,
      visibleFilters: ["marketFilter", "regionFilter", "jobFilter"].filter((id) => {
        const rect = document.getElementById(id).getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      }).length,
    }));
    assert(
      mobileLayout.documentWidth <= mobileLayout.viewportWidth + 1,
      `移动端发生横向溢出：${JSON.stringify(mobileLayout)}`,
    );
    assert.strictEqual(mobileLayout.visibleFilters, 3);
    assert((await paintedPixels(mobilePage, "#trendChart")) > 200);
    await mobilePage.screenshot({
      path: "/private/tmp/supplier-dashboard-americas-apac-mobile.png",
      fullPage: true,
    });
    assert.deepStrictEqual(
      mobileErrors,
      [],
      `移动端页面错误：${mobileErrors.join(" | ")}`,
    );
    await mobileContext.close();
    console.log(
      JSON.stringify(
        {
          status: "PASS",
          combined: { files: 7, markets: 2, suppliers: 60, output: combinedOutput },
          global: { files: 14, markets: 3, suppliers: 104, output: globalOutput },
          browser_errors: errors.length,
          mobile_errors: mobileErrors.length,
          mobile_layout: mobileLayout,
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

"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { pathToFileURL } = require("url");
const { chromium } = require("playwright");
const { launchOptions } = require("./playwright-launch.js");

const appDir = path.resolve(__dirname, "..");
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
const appUrl = process.env.APP_URL || pathToFileURL(path.join(appDir, "index.html")).href;

async function main() {
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
    await page.goto(appUrl, { waitUntil: "load" });
    await page.setInputFiles("#fileInput", sourceFiles);
    await page.click("#analyzeButton");
    await page.waitForSelector("#statusPanel.is-success", { timeout: 60000 });

    const downloadPromise = page.waitForEvent("download", { timeout: 20000 });
    await page.click("#exportButton");
    let download;
    try {
      download = await downloadPromise;
    } catch (error) {
      const status = await page.textContent("#statusText");
      throw new Error(`未触发下载；页面状态：${status}; 原始错误：${error.message}`);
    }
    assert.strictEqual(
      download.suggestedFilename(),
      "欧洲供应商履约率报表_2026年4-6月_最终严谨版.xlsx",
    );
    assert.deepStrictEqual(errors, []);
    await context.close();
    console.log(JSON.stringify({ status: "PASS", app_url: appUrl, errors: 0 }, null, 2));
  } finally {
    await browser.close();
  }
}

main().catch((error) => {
  console.error(error.stack || error);
  process.exitCode = 1;
});

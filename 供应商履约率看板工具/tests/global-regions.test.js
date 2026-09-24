"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("../vendor/xlsx.full.min.js");
const engine = require("../fulfillment-engine.js");

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

function exactArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function loadSources(directory, prefix = "") {
  return fs
    .readdirSync(directory)
    .filter((name) => name.endsWith(".xlsx") && !name.startsWith("~$"))
    .sort()
    .map((name) => {
      const fullPath = path.join(directory, name);
      const data = fs.readFileSync(fullPath);
      return {
        name: `${prefix}${name}`,
        size: data.length,
        lastModified: fs.statSync(fullPath).mtimeMs,
        data: exactArrayBuffer(data),
      };
    });
}

function filterSource(source, allowedMarkets) {
  const workbook = XLSX.read(source.data, {
    type: "array",
    cellDates: true,
    raw: true,
  });
  const output = XLSX.utils.book_new();
  ["基本信息", "需求详情", "发单详情"].forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      raw: true,
      defval: null,
      blankrows: false,
    });
    const headerCount = sheetName === "基本信息" ? 2 : 1;
    const regionColumn =
      sheetName === "基本信息"
        ? 0
        : rows[0].findIndex((value) => String(value || "").trim() === "需求区域");
    assert(regionColumn >= 0, `${sheetName}缺少需求区域`);
    const filtered = rows.slice(0, headerCount).concat(
      rows.slice(headerCount).filter((row) => {
        const region = String(row[regionColumn] || "").trim();
        return allowedMarkets.has(engine.marketForRegion(region));
      }),
    );
    XLSX.utils.book_append_sheet(output, XLSX.utils.aoa_to_sheet(filtered), sheetName);
  });
  const data = XLSX.write(output, { type: "array", bookType: "xlsx", cellDates: true });
  return {
    ...source,
    name: source.name.replace(/\.xlsx$/i, `_${Array.from(allowedMarkets).join("_")}.xlsx`),
    size: data.byteLength,
    data,
  };
}

function close(left, right, label, tolerance = 1e-6) {
  assert(
    Math.abs(Number(left || 0) - Number(right || 0)) <= tolerance,
    `${label}: ${left} != ${right}`,
  );
}

function assertTotalsAdd(left, right, total, label) {
  Object.keys(total).forEach((metric) => {
    close(
      Number(left[metric] || 0) + Number(right[metric] || 0),
      total[metric],
      `${label}.${metric}`,
    );
  });
}

function assertQuality(analysis, label) {
  assert.strictEqual(analysis.quality.status, "PASS", `${label}质量状态`);
  assert.strictEqual(analysis.meta.coverage.complete, true, `${label}日期覆盖`);
  assert.deepStrictEqual(analysis.meta.unknown_regions, [], `${label}未识别区域`);
  analysis.quality.critical_checks.forEach((check) => {
    assert.strictEqual(check.status, "PASS", `${label}.${check.id}`);
  });
}

function marketMetric(analysis, market, month, metric) {
  return analysis.supplier_monthly
    .filter(
      (row) =>
        engine.marketForRegion(row.region) === market &&
        row.month === month,
    )
    .reduce((sum, row) => sum + Number(row[metric] || 0), 0);
}

async function main() {
  const europeSources = loadSources(europeDir, "欧洲_");
  const combinedSources = loadSources(americasApacDir, "美洲亚太_");
  const americasSources = combinedSources.map((source) =>
    filterSource(source, new Set(["美洲"])),
  );
  const apacSources = combinedSources.map((source) =>
    filterSource(source, new Set(["亚太"])),
  );

  const [europe, americas, apac, combined] = await Promise.all([
    engine.analyzeFiles(europeSources),
    engine.analyzeFiles(americasSources),
    engine.analyzeFiles(apacSources),
    engine.analyzeFiles(combinedSources),
  ]);
  const global = await engine.analyzeFiles(europeSources.concat(combinedSources));

  assert.deepStrictEqual(combined.meta.markets, ["美洲", "亚太"]);
  assert.strictEqual(combined.meta.scope_prefix, "美洲亚太");
  assert.strictEqual(
    combined.meta.export_filename,
    "美洲亚太供应商履约率报表_2026年4-6月_最终严谨版.xlsx",
  );
  assert.deepStrictEqual(americas.meta.markets, ["美洲"]);
  assert.strictEqual(americas.meta.scope_prefix, "美洲");
  assert.deepStrictEqual(apac.meta.markets, ["亚太"]);
  assert.strictEqual(apac.meta.scope_prefix, "澳洲");
  assert.deepStrictEqual(global.meta.markets, ["欧洲", "美洲", "亚太"]);
  assert.strictEqual(global.meta.scope_prefix, "全球");
  assert.strictEqual(global.meta.scope_total_label, "全球总计");

  [
    ["欧洲", europe],
    ["美洲", americas],
    ["亚太", apac],
    ["美洲亚太", combined],
    ["全球", global],
  ].forEach(([label, analysis]) => assertQuality(analysis, label));

  assertTotalsAdd(americas.clean_totals, apac.clean_totals, combined.clean_totals, "美洲+亚太");
  assertTotalsAdd(europe.clean_totals, combined.clean_totals, global.clean_totals, "欧洲+美洲亚太");
  assert.strictEqual(combined.meta.supplier_count, 60);
  assert.strictEqual(global.meta.supplier_count, europe.meta.supplier_count + 60);

  combined.meta.months.forEach((month) => {
    ["美洲", "亚太"].forEach((market) => {
      const display = combined.display_metrics.find(
        (row) =>
          row.row_type === "market_total" &&
          row.market === market,
      );
      const period = display.by_month.find((row) => row.month === month.key);
      close(
        period.issued,
        marketMetric(combined, market, month.key, "issued"),
        `${market}.${month.key}.issued`,
      );
      close(
        period.raw_effective,
        marketMetric(combined, market, month.key, "raw_effective"),
        `${market}.${month.key}.raw_effective`,
      );
      close(
        period.reasonable_responsibility,
        marketMetric(combined, market, month.key, "reasonable_responsibility"),
        `${market}.${month.key}.responsibility`,
      );
    });
  });

  const globalTotal = global.display_metrics.find((row) => row.row_type === "scope_total");
  close(globalTotal.totals.issued, global.clean_totals.issued, "全球总计发单");
  close(globalTotal.totals.raw_effective, global.clean_totals.raw_effective, "全球总计兑现");
  close(
    globalTotal.totals.reasonable_responsibility,
    global.clean_totals.reasonable_responsibility,
    "全球总计合理责任",
  );

  console.log(
    JSON.stringify(
      {
        status: "PASS",
        scopes: {
          europe: { regions: europe.meta.regions.length, suppliers: europe.meta.supplier_count },
          americas: {
            regions: americas.meta.regions.length,
            suppliers: americas.meta.supplier_count,
          },
          apac: { regions: apac.meta.regions.length, suppliers: apac.meta.supplier_count },
          combined: {
            regions: combined.meta.regions.length,
            suppliers: combined.meta.supplier_count,
          },
          global: { regions: global.meta.regions.length, suppliers: global.meta.supplier_count },
        },
        combined_totals: combined.clean_totals,
        global_totals: global.clean_totals,
        quality_checks: global.quality.critical_checks,
      },
      null,
      2,
    ),
  );
}

main().catch((error) => {
  console.error(error.stack || error);
  process.exitCode = 1;
});

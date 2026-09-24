"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("../vendor/xlsx.full.min.js");
const engine = require("../fulfillment-engine.js");
const exporter = require("../report-exporter.js");

const root = path.resolve(__dirname, "../../../..");
const sourceDir = path.join(root, "需求满足率统计", "欧洲4月-6月需求满足率");
const templatePath = path.join(__dirname, "..", "vendor", "performance-report-template.xlsx");

function exactArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function allSources() {
  return fs
    .readdirSync(sourceDir)
    .filter((name) => name.endsWith(".xlsx") && !name.startsWith("~$"))
    .sort()
    .map((name) => {
      const fullPath = path.join(sourceDir, name);
      const data = fs.readFileSync(fullPath);
      return {
        name,
        size: data.length,
        lastModified: fs.statSync(fullPath).mtimeMs,
        data: exactArrayBuffer(data),
      };
    });
}

function iso(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(
    date.getDate(),
  ).padStart(2, "0")}`;
}

function shiftDate(date, months) {
  const absoluteMonth = date.getFullYear() * 12 + date.getMonth() + months;
  const year = Math.floor(absoluteMonth / 12);
  const month = absoluteMonth % 12;
  return new Date(year, month, date.getDate(), 12, 0, 0, 0);
}

function shiftBasicHeader(value, months) {
  let date = null;
  if (value instanceof Date) date = value;
  else if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) date = new Date(parsed.y, parsed.m - 1, parsed.d, 12, 0, 0, 0);
  } else {
    const match = String(value || "").trim().match(/^(\d{1,2})-(\d{1,2})/);
    if (match) date = new Date(2026, Number(match[1]) - 1, Number(match[2]), 12, 0, 0, 0);
  }
  if (!date) return value;
  const shifted = shiftDate(date, months);
  return `${shifted.getMonth() + 1}-${shifted.getDate()}`;
}

function shiftDetailDate(value, months) {
  const span = engine.parseSpan(value, 2026);
  if (!span) return value;
  const start = shiftDate(span[0], months);
  const end = shiftDate(span[1], months);
  return iso(start) === iso(end) ? iso(start) : `${iso(start)}~${iso(end)}`;
}

function shiftedSource(source, months) {
  if (months === 0) return source;
  const workbook = XLSX.read(source.data, { type: "array", cellDates: true, raw: true });
  const output = XLSX.utils.book_new();
  ["基本信息", "需求详情", "发单详情"].forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      raw: true,
      defval: null,
      blankrows: false,
    });
    if (sheetName === "基本信息") {
      for (let column = 3; column < (rows[0] || []).length; column += 1) {
        rows[0][column] = shiftBasicHeader(rows[0][column], months);
      }
    } else {
      const dateColumn = rows[0].findIndex(
        (value) => String(value || "").trim() === "需求日期",
      );
      assert(dateColumn >= 0, `${sheetName}缺少需求日期`);
      for (let row = 1; row < rows.length; row += 1) {
        rows[row][dateColumn] = shiftDetailDate(rows[row][dateColumn], months);
      }
    }
    XLSX.utils.book_append_sheet(output, XLSX.utils.aoa_to_sheet(rows), sheetName);
  });
  const data = XLSX.write(output, { type: "array", bookType: "xlsx", cellDates: true });
  return {
    name: source.name.replace(/\.xlsx$/i, `_shift_${months}.xlsx`),
    size: data.byteLength,
    lastModified: source.lastModified,
    data,
  };
}

async function exportAndInspect(analysis, outputPath, expectedRawEndColumn, expectedAdjustedEndColumn) {
  const template = new Uint8Array(fs.readFileSync(templatePath));
  const report = await exporter.buildReport(analysis, template);
  fs.writeFileSync(outputPath, Buffer.from(await report.blob.arrayBuffer()));
  const workbook = XLSX.read(fs.readFileSync(outputPath), { type: "buffer", cellFormula: true });
  const rawHeaders = XLSX.utils.sheet_to_json(workbook.Sheets["主指标_原始兑现率"], {
    header: 1,
    raw: true,
    defval: null,
  })[3];
  const adjustedHeaders = XLSX.utils.sheet_to_json(workbook.Sheets["对照_超发校正率"], {
    header: 1,
    raw: true,
    defval: null,
  })[3];
  assert.strictEqual(rawHeaders.length, XLSX.utils.decode_col(expectedRawEndColumn) + 1);
  assert.strictEqual(adjustedHeaders.length, XLSX.utils.decode_col(expectedAdjustedEndColumn) + 1);
  assert.strictEqual(
    rawHeaders.filter((header) => /履约率$/.test(String(header || ""))).length,
    analysis.meta.month_count,
  );
  return report.fileName;
}

async function main() {
  const sources = allSources();
  const aprilSources = sources.filter((source) => /\(4\./.test(source.name));
  const aprilMaySources = sources.filter((source) => /\([45]\./.test(source.name));
  const april = await engine.analyzeFiles(aprilSources);
  assert.deepStrictEqual(april.meta.months.map((month) => month.key), ["2026-04"]);
  assert.strictEqual(april.meta.period_label, "2026年4月");
  assert.strictEqual(april.meta.period_short, "统计期");
  assert.strictEqual(april.meta.coverage.complete, true);
  assert.strictEqual(april.quality.status, "PASS");
  const aprilOutput = "/private/tmp/供应商履约率报表_单月动态回归.xlsx";
  await exportAndInspect(april, aprilOutput, "J", "K");

  const aprilMay = await engine.analyzeFiles(aprilMaySources);
  assert.deepStrictEqual(aprilMay.meta.months.map((month) => month.key), [
    "2026-04",
    "2026-05",
  ]);
  assert.strictEqual(aprilMay.meta.period_label, "2026年4-5月");
  assert.strictEqual(aprilMay.meta.coverage.complete, true);
  assert.strictEqual(aprilMay.quality.status, "PASS");
  const twoMonthOutput = "/private/tmp/供应商履约率报表_双月动态回归.xlsx";
  await exportAndInspect(aprilMay, twoMonthOutput, "K", "L");

  const yearSources = [0, 3, 6, 9].flatMap((offset) =>
    sources.map((source) => shiftedSource(source, offset)),
  );
  if (process.env.FLEX_FIXTURE_DIR) {
    fs.mkdirSync(process.env.FLEX_FIXTURE_DIR, { recursive: true });
    yearSources.forEach((source) => {
      fs.writeFileSync(
        path.join(process.env.FLEX_FIXTURE_DIR, source.name),
        Buffer.from(source.data),
      );
    });
  }
  const year = await engine.analyzeFiles(yearSources);
  assert.strictEqual(year.meta.month_count, 12);
  assert.strictEqual(year.meta.months[0].key, "2026-04");
  assert.strictEqual(year.meta.months[11].key, "2027-03");
  assert.strictEqual(year.meta.period_short, "统计期");
  assert.strictEqual(
    year.quality.status,
    "PASS",
    JSON.stringify(year.quality.critical_checks),
  );
  const yearOutput = "/private/tmp/供应商履约率报表_12个月动态回归.xlsx";
  await exportAndInspect(year, yearOutput, "U", "V");

  console.log(
    JSON.stringify(
      {
        status: "PASS",
        periods: [
          {
            label: april.meta.period_label,
            months: april.meta.month_count,
            output: aprilOutput,
          },
          {
            label: aprilMay.meta.period_label,
            months: aprilMay.meta.month_count,
            output: twoMonthOutput,
          },
          {
            label: year.meta.period_label,
            months: year.meta.month_count,
            output: yearOutput,
            coverage_complete: year.meta.coverage.complete,
          },
        ],
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

"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const XLSX = require("../vendor/xlsx.full.min.js");
const engine = require("../fulfillment-engine.js");

const outputDir = process.env.SYNTHETIC_FIXTURE_DIR || "/private/tmp/supplier-dashboard-global-synthetic";
const cases = [
  {
    region: "英国区",
    warehouse: "虚构欧洲仓",
    group: "虚构操作组",
    job: "操作员",
    supplier: "Synthetic Europe Supplier",
    supplierId: "SYN-EU",
    demand: 10,
    issued: 12,
    dispatched: 8,
  },
  {
    region: "达拉斯区",
    warehouse: "虚构美洲仓",
    group: "虚构叉车组",
    job: "叉车司机",
    supplier: "Synthetic Americas Supplier",
    supplierId: "SYN-AM",
    demand: 20,
    issued: 15,
    dispatched: 12,
  },
  {
    region: "加拿大区",
    warehouse: "虚构美洲北区仓",
    group: "虚构包装组",
    job: "包装员",
    supplier: "Synthetic Americas North Supplier",
    supplierId: "SYN-AM-N",
    demand: 6,
    issued: 9,
    dispatched: 5,
  },
  {
    region: "澳洲区",
    warehouse: "虚构亚太仓",
    group: "虚构分拣组",
    job: "分拣员",
    supplier: "Synthetic APAC Supplier",
    supplierId: "SYN-AP",
    demand: 8,
    issued: 8,
    dispatched: 7,
  },
];

function exactArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function workbookForMonth(month) {
  const date = `2026-${String(month).padStart(2, "0")}-15`;
  const basicMetrics = [
    "需求总人数",
    "补员",
    "待发单人数",
    "已发单人数",
    "待派遣人数",
    "已派遣人数",
    "需求完成率：百分比",
  ];
  const basic = [
    ["区域", "仓库", "组", `${month}-15`, null, null, null, null, null, null],
    [null, null, null, ...basicMetrics],
    ...cases.map((item) => [
      item.region,
      item.warehouse,
      item.group,
      item.demand,
      0,
      Math.max(item.demand - item.issued, 0),
      item.issued,
      Math.max(item.issued - item.dispatched, 0),
      item.dispatched,
      item.demand ? item.dispatched / item.demand : 0,
    ]),
  ];
  const demand = [
    [
      "需求区域",
      "需求仓",
      "需求组",
      "需求日期",
      "工种",
      "班次",
      "需求人数",
      "待发单人数",
      "已发单人数",
      "待派遣人数",
      "已派遣人数",
      "关联申请编号",
    ],
    ...cases.map((item) => [
      item.region,
      item.warehouse,
      item.group,
      date,
      item.job,
      "白班",
      item.demand,
      Math.max(item.demand - item.issued, 0),
      item.issued,
      Math.max(item.issued - item.dispatched, 0),
      item.dispatched,
      `SYN-${month}-${item.supplierId}`,
    ]),
  ];
  const dispatch = [
    [
      "需求区域",
      "需求仓",
      "需求组",
      "需求日期",
      "工种",
      "班次",
      "供应商名称",
      "供应商ID",
      "类型",
      "已发单人数",
      "待派遣人数",
      "已派遣人数",
    ],
    ...cases.map((item) => [
      item.region,
      item.warehouse,
      item.group,
      date,
      item.job,
      "白班",
      item.supplier,
      item.supplierId,
      "正式",
      item.issued,
      Math.max(item.issued - item.dispatched, 0),
      item.dispatched,
    ]),
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(basic), "基本信息");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(demand), "需求详情");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(dispatch), "发单详情");
  return Buffer.from(XLSX.write(workbook, { type: "array", bookType: "xlsx" }));
}

async function main() {
  fs.mkdirSync(outputDir, { recursive: true });
  const sources = [4, 5, 6].map((month) => {
    const name = `synthetic-global-2026-${String(month).padStart(2, "0")}.xlsx`;
    const buffer = workbookForMonth(month);
    fs.writeFileSync(path.join(outputDir, name), buffer);
    return {
      name,
      size: buffer.length,
      lastModified: Date.UTC(2026, month - 1, 15),
      data: exactArrayBuffer(buffer),
    };
  });
  const analysis = await engine.analyzeFiles(sources);
  assert.strictEqual(analysis.quality.status, "PASS");
  assert.deepStrictEqual(analysis.meta.markets, ["欧洲", "美洲", "亚太"]);
  assert.strictEqual(analysis.meta.regions.length, 4);
  assert.strictEqual(analysis.meta.supplier_count, 4);
  console.log(
    JSON.stringify(
      {
        status: "PASS",
        output_dir: outputDir,
        files: sources.length,
        markets: analysis.meta.markets,
        regions: analysis.meta.regions,
        suppliers: analysis.meta.supplier_count,
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

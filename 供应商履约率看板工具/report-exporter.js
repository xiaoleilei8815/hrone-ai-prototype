(function (root, factory) {
  const api = factory(root);
  if (typeof module === "object" && module.exports) module.exports = api;
  root.SupplierFulfillmentReportExporter = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function (root) {
  "use strict";

  const TEMPLATE_PATH = "vendor/performance-report-template.xlsx";
  const TABLE_NAMES = [
    "RawFulfillmentScorecard",
    "AdjustedFulfillmentScorecard",
    "SupplierMonthlyMetrics",
    "WarehouseAndAllocationSummary",
    "SupplierAtomicResponsibility",
    "WarehouseAtomicFacts",
    "MonthlyThreeWayReconciliation",
    "SourceFileRegistry",
  ];

  function getXlsxPopulate() {
    if (root.XlsxPopulate) return root.XlsxPopulate;
    if (typeof require === "function") {
      return require("./vendor/xlsx-populate-no-encryption.min.js");
    }
    throw new Error("履约率报表模板引擎未加载");
  }

  function xmlEscape(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function localDate(iso) {
    if (!iso) return null;
    const [year, month, day] = String(iso).slice(0, 10).split("-").map(Number);
    return new Date(year, month - 1, day, 0, 0, 0, 0);
  }

  function number(value) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function cachedValue(value) {
    if (value == null || value === "") return "";
    if (typeof value !== "number") return value;
    return Number.isFinite(value) ? Number(value.toPrecision(15)) : "";
  }

  function setFormula(cell, formula, value) {
    const normalized = String(formula || "").replace(/^=/, "");
    const cached = cachedValue(value);
    cell.formula(normalized);
    const attributes = { ...(cell._remainingAttributes || {}) };
    delete attributes.t;
    if (typeof cached === "string") attributes.t = "str";
    cell._remainingAttributes = attributes;
    cell._remainingChildren = [
      { name: "v", children: [cached == null ? "" : cached] },
    ];
    return cell;
  }

  function ratioFormula(numerator, denominator) {
    return `IF((${denominator})=0,"",(${numerator})/(${denominator}))`;
  }

  function columnName(numberValue) {
    let value = numberValue;
    let result = "";
    while (value > 0) {
      value -= 1;
      result = String.fromCharCode(65 + (value % 26)) + result;
      value = Math.floor(value / 26);
    }
    return result;
  }

  function captureRowStyles(sheet, row, columns) {
    return Array.from({ length: columns }, (_, index) => {
      return sheet.cell(row, index + 1)._styleId;
    });
  }

  function applyRowStyles(sheet, row, styles) {
    styles.forEach((styleId, index) => {
      sheet.cell(row, index + 1)._styleId = styleId;
    });
  }

  function fillRows(sheet, startRow, rows, styleResolver) {
    if (!rows.length) return startRow - 1;
    const columns = rows[0].length;
    sheet.range(startRow, 1, startRow + rows.length - 1, columns).value(rows);
    rows.forEach((_row, index) => {
      applyRowStyles(sheet, startRow + index, styleResolver(index));
    });
    return startRow + rows.length - 1;
  }

  function trimSheet(sheet, rowCount) {
    if (sheet._rows.length > rowCount + 1) {
      sheet._rows.length = rowCount + 1;
      sheet._sheetDataNode.children = sheet._rows;
    }
  }

  function clearValues(sheet) {
    const used = sheet.usedRange();
    if (used) used.value(null);
  }

  function monthHeaders(analysis) {
    return analysis.meta.months.map((month) => `${month.label}履约率`);
  }

  function metricForMonth(row, monthKey) {
    return row.by_month.find((item) => item.month === monthKey) || {};
  }

  function displayRowStyleSource(rowType) {
    const sources = {
      supplier_job: 5,
      supplier_total: 6,
      region_job: 31,
      region_total: 37,
      europe_job: 148,
      europe_total: 155,
      market_job: 148,
      market_total: 155,
      scope_job: 148,
      scope_total: 155,
    };
    return sources[rowType] || 5;
  }

  function warehouseRowStyleSource(rowType) {
    const sources = {
      region_job: 5,
      region_total: 15,
      europe_job: 60,
      europe_total: 77,
      market_job: 60,
      market_total: 77,
      scope_job: 60,
      scope_total: 77,
    };
    return sources[rowType] || 5;
  }

  function sumExpression(sheetName, sumColumn, startRow, endRow, criteria) {
    const sumRange = `'${sheetName}'!$${sumColumn}$${startRow}:$${sumColumn}$${endRow}`;
    if (!criteria.length) return `SUM(${sumRange})`;
    return `SUMIFS(${sumRange},${criteria.join(",")})`;
  }

  function excelText(value) {
    return `"${String(value == null ? "" : value).replace(/"/g, '""')}"`;
  }

  function sumForDisplayRegions(
    sheetName,
    sumColumn,
    startRow,
    endRow,
    displayRow,
    regionRange,
    criteria,
    regionCell,
  ) {
    if (["market_job", "market_total"].includes(displayRow.row_type)) {
      const regions = displayRow.member_regions || [];
      if (!regions.length) return "0";
      return `(${regions
        .map((region) =>
          sumExpression(sheetName, sumColumn, startRow, endRow, [
            regionRange,
            excelText(region),
            ...criteria,
          ]),
        )
        .join("+")})`;
    }
    if (
      ["supplier_job", "supplier_total", "region_job", "region_total"].includes(
        displayRow.row_type,
      )
    ) {
      criteria.unshift(regionRange, regionCell);
    }
    return sumExpression(sheetName, sumColumn, startRow, endRow, criteria);
  }

  function supplierMonthlyExpression(
    analysis,
    sumColumn,
    displayRow,
    excelRow,
    monthlyStart,
    monthlyEnd,
    month,
  ) {
    const criteria = [];
    if (["supplier_job", "supplier_total"].includes(displayRow.row_type)) {
      criteria.push(
        `'供应商月度明细'!$C$${monthlyStart}:$C$${monthlyEnd}`,
        `$C${excelRow}`,
      );
    }
    if (
      ["supplier_job", "region_job", "market_job", "scope_job", "europe_job"].includes(
        displayRow.row_type,
      )
    ) {
      criteria.push(
        `'供应商月度明细'!$D$${monthlyStart}:$D$${monthlyEnd}`,
        `$E${excelRow}`,
      );
    }
    if (month) {
      criteria.push(
        `'供应商月度明细'!$E$${monthlyStart}:$E$${monthlyEnd}`,
        `"${month}"`,
      );
    }
    return sumForDisplayRegions(
      "供应商月度明细",
      sumColumn,
      monthlyStart,
      monthlyEnd,
      displayRow,
      `'供应商月度明细'!$A$${monthlyStart}:$A$${monthlyEnd}`,
      criteria,
      `$A${excelRow}`,
    );
  }

  function unitExpression(sumColumn, row, excelRow, unitStart, unitEnd) {
    const criteria = [];
    if (["region_job", "market_job", "scope_job", "europe_job"].includes(row.row_type)) {
      criteria.push(
        `'仓库单元底表'!$F$${unitStart}:$F$${unitEnd}`,
        `$C${excelRow}`,
      );
    }
    criteria.push(
      `'仓库单元底表'!$E$${unitStart}:$E$${unitEnd}`,
      `$D${excelRow}`,
    );
    return sumForDisplayRegions(
      "仓库单元底表",
      sumColumn,
      unitStart,
      unitEnd,
      row,
      `'仓库单元底表'!$A$${unitStart}:$A$${unitEnd}`,
      criteria,
      `$A${excelRow}`,
    );
  }

  function supplierSummaryExpression(
    sumColumn,
    row,
    excelRow,
    monthlyStart,
    monthlyEnd,
  ) {
    const criteria = [];
    if (["region_job", "market_job", "scope_job", "europe_job"].includes(row.row_type)) {
      criteria.push(
        `'供应商月度明细'!$D$${monthlyStart}:$D$${monthlyEnd}`,
        `$C${excelRow}`,
      );
    }
    criteria.push(
      `'供应商月度明细'!$E$${monthlyStart}:$E$${monthlyEnd}`,
      `$D${excelRow}`,
    );
    return sumForDisplayRegions(
      "供应商月度明细",
      sumColumn,
      monthlyStart,
      monthlyEnd,
      row,
      `'供应商月度明细'!$A$${monthlyStart}:$A$${monthlyEnd}`,
      criteria,
      `$A${excelRow}`,
    );
  }

  function setSupplierAtomic(workbook, analysis) {
    const sheet = workbook.sheet("供应商责任底表");
    const headerStyles = captureRowStyles(sheet, 4, 22);
    const bodyStyles = captureRowStyles(sheet, 5, 22);
    clearValues(sheet);
    sheet.cell("A1").value("供应商责任底表：最细单元的实际订单与合理责任");
    sheet
      .cell("A2")
      .value(
        "事实字段保持真实人日；合理责任与校正兑现为按发单占比折算的责任人日等值。所有汇总表均由本页公式链向上聚合。",
      );
    const headers = [
      "区域",
      "仓库",
      "组",
      "日期",
      "月份",
      "工种",
      "班次",
      "供应商ID",
      "供应商名称",
      "原始需求",
      "补员",
      "真实需求",
      "单元总发单",
      "单元总派遣",
      "供应商发单",
      "供应商派遣",
      "原始有效兑现",
      "责任折算因子",
      "合理责任",
      "校正有效兑现",
      "释放超发责任",
      "派遣超订单",
    ];
    sheet.range("A4:V4").value([headers]);
    applyRowStyles(sheet, 4, headerStyles);
    const unitMap = new Map(
      analysis.unit_atomic.map((row) => [
        [row.region, row.warehouse, row.group, row.date, row.job, row.shift].join(
          "\u0001",
        ),
        row,
      ]),
    );
    const start = 5;
    const values = analysis.supplier_atomic.map((row) => {
      const unit = unitMap.get(
        [row.region, row.warehouse, row.group, row.date, row.job, row.shift].join(
          "\u0001",
        ),
      );
      return [
        row.region,
        row.warehouse,
        row.group,
        localDate(row.date),
        row.month,
        row.job,
        row.shift,
        row.supplier_id,
        row.supplier_name,
        number(unit.raw_demand),
        number(unit.supplement),
        null,
        number(row.unit_total_issued),
        number(row.unit_total_dispatched),
        number(row.issued),
        number(row.dispatched),
        null,
        null,
        null,
        null,
        null,
        null,
      ];
    });
    const end = fillRows(sheet, start, values, () => bodyStyles);
    analysis.supplier_atomic.forEach((row, index) => {
      const excelRow = start + index;
      const unit = unitMap.get(
        [row.region, row.warehouse, row.group, row.date, row.job, row.shift].join(
          "\u0001",
        ),
      );
      setFormula(sheet.cell(excelRow, 12), `J${excelRow}+K${excelRow}`, unit.true_demand);
      setFormula(
        sheet.cell(excelRow, 17),
        `MIN(O${excelRow},P${excelRow})`,
        row.raw_effective,
      );
      setFormula(
        sheet.cell(excelRow, 18),
        `IF(M${excelRow}=0,0,MIN(1,L${excelRow}/M${excelRow}))`,
        row.responsibility_factor,
      );
      setFormula(
        sheet.cell(excelRow, 19),
        `O${excelRow}*R${excelRow}`,
        row.reasonable_responsibility,
      );
      setFormula(
        sheet.cell(excelRow, 20),
        `MIN(P${excelRow},S${excelRow})`,
        row.adjusted_effective,
      );
      setFormula(
        sheet.cell(excelRow, 21),
        `O${excelRow}-S${excelRow}`,
        row.over_issue_relief,
      );
      setFormula(
        sheet.cell(excelRow, 22),
        `MAX(P${excelRow}-O${excelRow},0)`,
        row.dispatch_over_order,
      );
    });
    trimSheet(sheet, end);
    return { start, end, headers };
  }

  function setUnitAtomic(workbook, analysis) {
    const sheet = workbook.sheet("仓库单元底表");
    const headerStyles = captureRowStyles(sheet, 4, 19);
    const bodyStyles = captureRowStyles(sheet, 5, 19);
    clearValues(sheet);
    sheet.cell("A1").value("仓库单元底表：需求、发单、派遣与HR分配质量");
    sheet
      .cell("A2")
      .value(
        "原子颗粒度为区域×仓库×组×日期×工种×班次。需求满足和超发均先在本页逐单元封顶或计算，再向上汇总。",
      );
    const headers = [
      "区域",
      "仓库",
      "组",
      "日期",
      "月份",
      "工种",
      "班次",
      "原始需求",
      "补员",
      "真实需求",
      "总发单",
      "总派遣",
      "发单覆盖量",
      "仓库有效满足",
      "责任折算因子",
      "超发人日",
      "零需求发单",
      "少发人日",
      "冗余派遣",
    ];
    sheet.range("A4:S4").value([headers]);
    applyRowStyles(sheet, 4, headerStyles);
    const start = 5;
    const values = analysis.unit_atomic.map((row) => [
      row.region,
      row.warehouse,
      row.group,
      localDate(row.date),
      row.month,
      row.job,
      row.shift,
      number(row.raw_demand),
      number(row.supplement),
      null,
      number(row.issued),
      number(row.dispatched),
      null,
      null,
      null,
      null,
      null,
      null,
      null,
    ]);
    const end = fillRows(sheet, start, values, () => bodyStyles);
    analysis.unit_atomic.forEach((row, index) => {
      const excelRow = start + index;
      setFormula(sheet.cell(excelRow, 10), `H${excelRow}+I${excelRow}`, row.true_demand);
      setFormula(
        sheet.cell(excelRow, 13),
        `MIN(K${excelRow},J${excelRow})`,
        row.issue_covered,
      );
      setFormula(
        sheet.cell(excelRow, 14),
        `MIN(L${excelRow},J${excelRow})`,
        row.warehouse_effective,
      );
      setFormula(
        sheet.cell(excelRow, 15),
        `IF(K${excelRow}=0,0,MIN(1,J${excelRow}/K${excelRow}))`,
        row.responsibility_factor,
      );
      setFormula(
        sheet.cell(excelRow, 16),
        `IF(J${excelRow}>0,MAX(K${excelRow}-J${excelRow},0),0)`,
        row.over_issue,
      );
      setFormula(
        sheet.cell(excelRow, 17),
        `IF(J${excelRow}=0,K${excelRow},0)`,
        row.zero_demand_issue,
      );
      setFormula(
        sheet.cell(excelRow, 18),
        `MAX(J${excelRow}-K${excelRow},0)`,
        row.under_issue,
      );
      setFormula(
        sheet.cell(excelRow, 19),
        `IF(J${excelRow}>0,MAX(L${excelRow}-J${excelRow},0),L${excelRow})`,
        row.surplus_dispatch,
      );
    });
    trimSheet(sheet, end);
    return { start, end, headers };
  }

  function setSupplierMonthly(workbook, analysis, atomicRange) {
    const sheet = workbook.sheet("供应商月度明细");
    const headerStyles = captureRowStyles(sheet, 4, 15);
    const bodyStyles = captureRowStyles(sheet, 5, 15);
    clearValues(sheet);
    sheet.cell("A1").value("供应商×工种×月份：两套履约指标计算明细");
    sheet
      .cell("A2")
      .value(
        `原始接单兑现率为主指标；超发校正履约率为对照。分母为0时显示空白，不参与${
          analysis.meta.month_count === 3 && /^Q\d$/.test(analysis.meta.period_short)
            ? "三个月平均"
            : "统计期月均"
        }；有发单但未兑现显示0%。`,
      );
    const headers = [
      "区域",
      "供应商名称",
      "供应商ID",
      "工种",
      "月份",
      "实际发单",
      "实际派遣",
      "原始有效兑现",
      "原始兑现率",
      "合理责任",
      "校正有效兑现",
      "超发校正率",
      "释放超发责任",
      "校正率变化",
      "派遣超订单",
    ];
    sheet.range("A4:O4").value([headers]);
    applyRowStyles(sheet, 4, headerStyles);
    const start = 5;
    const values = analysis.supplier_monthly.map((row) => [
      row.region,
      row.supplier_name,
      row.supplier_id,
      row.job,
      row.month,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
    ]);
    const end = fillRows(sheet, start, values, () => bodyStyles);

    function atomicSum(column, excelRow) {
      return `SUMIFS('供应商责任底表'!$${column}$${atomicRange.start}:$${column}$${atomicRange.end},'供应商责任底表'!$A$${atomicRange.start}:$A$${atomicRange.end},$A${excelRow},'供应商责任底表'!$H$${atomicRange.start}:$H$${atomicRange.end},$C${excelRow},'供应商责任底表'!$F$${atomicRange.start}:$F$${atomicRange.end},$D${excelRow},'供应商责任底表'!$E$${atomicRange.start}:$E$${atomicRange.end},$E${excelRow})`;
    }

    analysis.supplier_monthly.forEach((row, index) => {
      const excelRow = start + index;
      setFormula(sheet.cell(excelRow, 6), atomicSum("O", excelRow), row.issued);
      setFormula(sheet.cell(excelRow, 7), atomicSum("P", excelRow), row.dispatched);
      setFormula(sheet.cell(excelRow, 8), atomicSum("Q", excelRow), row.raw_effective);
      setFormula(
        sheet.cell(excelRow, 9),
        ratioFormula(`H${excelRow}`, `F${excelRow}`),
        row.raw_rate,
      );
      setFormula(
        sheet.cell(excelRow, 10),
        atomicSum("S", excelRow),
        row.reasonable_responsibility,
      );
      setFormula(
        sheet.cell(excelRow, 11),
        atomicSum("T", excelRow),
        row.adjusted_effective,
      );
      setFormula(
        sheet.cell(excelRow, 12),
        ratioFormula(`K${excelRow}`, `J${excelRow}`),
        row.adjusted_rate,
      );
      setFormula(
        sheet.cell(excelRow, 13),
        atomicSum("U", excelRow),
        row.over_issue_relief,
      );
      setFormula(
        sheet.cell(excelRow, 14),
        `IF(OR(F${excelRow}=0,J${excelRow}=0),"",L${excelRow}-I${excelRow})`,
        row.rate_lift,
      );
      setFormula(
        sheet.cell(excelRow, 15),
        atomicSum("V", excelRow),
        row.dispatch_over_order,
      );
    });
    trimSheet(sheet, end);
    return { start, end, headers };
  }

  function expandMainStyles(templateStyles, monthCount, isRaw) {
    const trailingCount = isRaw ? 4 : 5;
    return templateStyles
      .slice(0, 5)
      .concat(Array.from({ length: monthCount }, () => templateStyles[5]))
      .concat(templateStyles.slice(8, 8 + trailingCount));
  }

  function setMainSheet(workbook, analysis, metric, monthlyRange) {
    const isRaw = metric === "raw";
    const sheetName = isRaw ? "主指标_原始兑现率" : "对照_超发校正率";
    const sheet = workbook.sheet(sheetName);
    const templateColumns = isRaw ? 12 : 13;
    const monthCount = analysis.meta.months.length;
    const columns = 5 + monthCount + 4 + (isRaw ? 0 : 1);
    const headerStyles = expandMainStyles(
      captureRowStyles(sheet, 4, templateColumns),
      monthCount,
      isRaw,
    );
    const rowStyles = {};
    [
      "supplier_job",
      "supplier_total",
      "region_job",
      "region_total",
      "europe_job",
      "europe_total",
      "market_job",
      "market_total",
      "scope_job",
      "scope_total",
    ].forEach((rowType) => {
      rowStyles[rowType] = expandMainStyles(
        captureRowStyles(
          sheet,
          displayRowStyleSource(rowType),
          templateColumns,
        ),
        monthCount,
        isRaw,
      );
    });
    const templateEndColumn = columnName(templateColumns);
    [1, 2].forEach((row) => {
      const range = sheet.range(`A${row}:${templateEndColumn}${row}`);
      if (range.merged()) range.merged(false);
    });
    clearValues(sheet);
    sheet
      .cell("A1")
      .value(
        `${analysis.meta.period_label}${analysis.meta.scope_prefix}供应商履约率：${
          isRaw ? "原始接单兑现率（主指标）" : "超发校正履约率（对照）"
        }`,
      );
    sheet
      .cell("A2")
      .value(
        isRaw
          ? "主指标 = Σmin(供应商派遣,供应商发单) ÷ Σ供应商发单。保留HR超发对实际订单兑现的影响，用于供应商盘点评分；月度无有效发单显示“—”。"
          : "对照指标 = Σmin(供应商派遣,合理责任) ÷ Σ合理责任；合理责任=发单×min(1,真实需求/单元总发单)。仅在超发单元缩减责任，结果为责任人日等值。",
      );
    const headers = [
      "区域",
      "供应商名称",
      "供应商ID",
      "层级",
      "工种",
      ...monthHeaders(analysis),
      "月度平均",
      `${analysis.meta.period_short}累计加权`,
      isRaw
        ? `${analysis.meta.period_short}实际发单`
        : `${analysis.meta.period_short}合理责任`,
      isRaw
        ? `${analysis.meta.period_short}有效兑现`
        : `${analysis.meta.period_short}校正兑现`,
    ];
    if (!isRaw) headers.push(`${analysis.meta.period_short}释放超发责任`);
    const endColumn = columnName(columns);
    sheet.range(`A1:${endColumn}1`).merged(true);
    sheet.range(`A2:${endColumn}2`).merged(true);
    sheet.range(`A4:${endColumn}4`).value([headers]);
    applyRowStyles(sheet, 4, headerStyles);
    const start = 5;
    const values = analysis.display_metrics.map((row) => {
      const base = Array.from({ length: columns }, () => null);
      [row.region, row.supplier_name, row.supplier_id, row.level, row.job].forEach(
        (value, index) => {
          base[index] = value;
        },
      );
      return base;
    });
    const end = fillRows(
      sheet,
      start,
      values,
      (index) => rowStyles[analysis.display_metrics[index].row_type],
    );
    analysis.display_metrics.forEach((row, index) => {
      const excelRow = start + index;
      const firstMonthColumn = 6;
      const lastMonthColumn = firstMonthColumn + monthCount - 1;
      const averageColumn = lastMonthColumn + 1;
      const weightedColumn = averageColumn + 1;
      const denominatorColumn = weightedColumn + 1;
      const numeratorColumn = denominatorColumn + 1;
      const reliefColumn = numeratorColumn + 1;
      analysis.meta.months.forEach((month, monthIndex) => {
        const denominator = supplierMonthlyExpression(
          analysis,
          isRaw ? "F" : "J",
          row,
          excelRow,
          monthlyRange.start,
          monthlyRange.end,
          month.key,
        );
        const numerator = supplierMonthlyExpression(
          analysis,
          isRaw ? "H" : "K",
          row,
          excelRow,
          monthlyRange.start,
          monthlyRange.end,
          month.key,
        );
        const monthMetric = metricForMonth(row, month.key);
        setFormula(
          sheet.cell(excelRow, 6 + monthIndex),
          ratioFormula(numerator, denominator),
          isRaw ? monthMetric.raw_rate : monthMetric.adjusted_rate,
        );
      });
      const firstMonthLetter = columnName(firstMonthColumn);
      const lastMonthLetter = columnName(lastMonthColumn);
      setFormula(
        sheet.cell(excelRow, averageColumn),
        `IF(COUNT(${firstMonthLetter}${excelRow}:${lastMonthLetter}${excelRow})=0,"",AVERAGE(${firstMonthLetter}${excelRow}:${lastMonthLetter}${excelRow}))`,
        isRaw ? row.raw_average : row.adjusted_average,
      );
      const totalDenominator = supplierMonthlyExpression(
        analysis,
        isRaw ? "F" : "J",
        row,
        excelRow,
        monthlyRange.start,
        monthlyRange.end,
      );
      const totalNumerator = supplierMonthlyExpression(
        analysis,
        isRaw ? "H" : "K",
        row,
        excelRow,
        monthlyRange.start,
        monthlyRange.end,
      );
      setFormula(
        sheet.cell(excelRow, denominatorColumn),
        totalDenominator,
        isRaw ? row.totals.issued : row.totals.reasonable_responsibility,
      );
      setFormula(
        sheet.cell(excelRow, numeratorColumn),
        totalNumerator,
        isRaw ? row.totals.raw_effective : row.totals.adjusted_effective,
      );
      const denominatorLetter = columnName(denominatorColumn);
      const numeratorLetter = columnName(numeratorColumn);
      setFormula(
        sheet.cell(excelRow, weightedColumn),
        ratioFormula(
          `${numeratorLetter}${excelRow}`,
          `${denominatorLetter}${excelRow}`,
        ),
        isRaw ? row.raw_weighted : row.adjusted_weighted,
      );
      if (!isRaw) {
        setFormula(
          sheet.cell(excelRow, reliefColumn),
          supplierMonthlyExpression(
            analysis,
            "M",
            row,
            excelRow,
            monthlyRange.start,
            monthlyRange.end,
          ),
          row.totals.over_issue_relief,
        );
      }
    });
    trimSheet(sheet, end);
    return {
      start,
      end,
      headers,
      averageColumn: 6 + monthCount,
      weightedColumn: 7 + monthCount,
      denominatorColumn: 8 + monthCount,
      numeratorColumn: 9 + monthCount,
      reliefColumn: isRaw ? null : 10 + monthCount,
    };
  }

  function setWarehouseSummary(workbook, analysis, unitRange, monthlyRange) {
    const sheet = workbook.sheet("仓库满足与发单质量");
    const headerStyles = captureRowStyles(sheet, 4, 17);
    const rowStyles = {};
    [
      "region_job",
      "region_total",
      "europe_job",
      "europe_total",
      "market_job",
      "market_total",
      "scope_job",
      "scope_total",
    ].forEach((rowType) => {
      rowStyles[rowType] = captureRowStyles(
        sheet,
        warehouseRowStyleSource(rowType),
        17,
      );
    });
    clearValues(sheet);
    sheet.cell("A1").value("仓库满足与发单质量：仓库结果、HR责任与供应商表现并列");
    sheet
      .cell("A2")
      .value(
        "仓库满足率只评价仓库保障结果；发单覆盖、超发、少发和零需求发单用于评价HR分配质量；原始/校正履约率用于观察供应商表现。",
      );
    const headers = [
      "区域",
      "层级",
      "工种",
      "月份",
      "真实需求",
      "总发单",
      "总派遣",
      "发单覆盖量",
      "仓库有效满足",
      "超发人日",
      "零需求发单",
      "少发人日",
      "冗余派遣",
      "发单覆盖率",
      "仓库满足率",
      "原始兑现率",
      "超发校正率",
    ];
    sheet.range("A4:Q4").value([headers]);
    applyRowStyles(sheet, 4, headerStyles);
    const start = 5;
    const values = analysis.warehouse_display.map((row) => [
      row.region,
      row.level,
      row.job,
      row.month,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
    ]);
    const end = fillRows(
      sheet,
      start,
      values,
      (index) => rowStyles[analysis.warehouse_display[index].row_type],
    );
    const unitColumns = ["J", "K", "L", "M", "N", "P", "Q", "R", "S"];
    const metricNames = [
      "true_demand",
      "issued",
      "dispatched",
      "issue_covered",
      "warehouse_effective",
      "over_issue",
      "zero_demand_issue",
      "under_issue",
      "surplus_dispatch",
    ];
    analysis.warehouse_display.forEach((row, index) => {
      const excelRow = start + index;
      unitColumns.forEach((column, metricIndex) => {
        setFormula(
          sheet.cell(excelRow, 5 + metricIndex),
          unitExpression(column, row, excelRow, unitRange.start, unitRange.end),
          row[metricNames[metricIndex]],
        );
      });
      setFormula(
        sheet.cell(excelRow, 14),
        ratioFormula(`H${excelRow}`, `E${excelRow}`),
        row.issue_coverage_rate,
      );
      setFormula(
        sheet.cell(excelRow, 15),
        ratioFormula(`I${excelRow}`, `E${excelRow}`),
        row.warehouse_fulfillment_rate,
      );
      const rawDen = supplierSummaryExpression(
        "F",
        row,
        excelRow,
        monthlyRange.start,
        monthlyRange.end,
      );
      const rawNum = supplierSummaryExpression(
        "H",
        row,
        excelRow,
        monthlyRange.start,
        monthlyRange.end,
      );
      const adjustedDen = supplierSummaryExpression(
        "J",
        row,
        excelRow,
        monthlyRange.start,
        monthlyRange.end,
      );
      const adjustedNum = supplierSummaryExpression(
        "K",
        row,
        excelRow,
        monthlyRange.start,
        monthlyRange.end,
      );
      setFormula(
        sheet.cell(excelRow, 16),
        ratioFormula(rawNum, rawDen),
        row.raw_rate,
      );
      setFormula(
        sheet.cell(excelRow, 17),
        ratioFormula(adjustedNum, adjustedDen),
        row.adjusted_rate,
      );
    });
    trimSheet(sheet, end);
    return { start, end, headers };
  }

  function setDashboard(workbook, analysis, ranges) {
    const sheet = workbook.sheet("管理摘要");
    const styles = {
      month: captureRowStyles(sheet, 6, 7),
      average: captureRowStyles(sheet, 9, 7),
      weighted: captureRowStyles(sheet, 10, 7),
      section: captureRowStyles(sheet, 12, 8),
      keyPrimary: captureRowStyles(sheet, 13, 8),
      keyRisk: captureRowStyles(sheet, 14, 8),
      keyRates: captureRowStyles(sheet, 15, 8),
      regionSection: captureRowStyles(sheet, 17, 8),
      regionHeader: captureRowStyles(sheet, 18, 8),
      regionBody: captureRowStyles(sheet, 19, 8),
      regionRightBlank: captureRowStyles(sheet, 19, 12).slice(8),
      conclusion: captureRowStyles(sheet, 27, 12),
      conclusionSecond: captureRowStyles(sheet, 28, 12),
    };
    const rowHeights = {
      regionHeader: sheet.row(18).height(),
      conclusion: sheet.row(27).height(),
      conclusionSecond: sheet.row(28).height(),
    };
    ["A12:H12", "A17:H17", "A27:L28"].forEach((address) => {
      const range = sheet.range(address);
      if (range.merged()) range.merged(false);
    });
    clearValues(sheet);

    const monthStart = 6;
    const monthEnd = monthStart + analysis.meta.months.length - 1;
    const averageRow = monthEnd + 1;
    const weightedRow = averageRow + 1;
    const keyTitleRow = weightedRow + 2;
    const keyPrimaryRow = keyTitleRow + 1;
    const keyRiskRow = keyTitleRow + 2;
    const keyRatesRow = keyTitleRow + 3;
    const regionTitleRow = keyTitleRow + 5;
    const regionHeaderRow = regionTitleRow + 1;
    const comparisonRows = analysis.comparison_summary || analysis.region_summary;
    const regionStartRow = regionHeaderRow + 1;
    const regionEndRow = regionStartRow + comparisonRows.length - 1;
    const conclusionRow = regionEndRow + 3;

    sheet.cell("A1").value(analysis.meta.management_title);
    sheet
      .cell("A2")
      .value(
        "主指标是原始接单兑现率；超发校正率用于判断HR超发是否稀释供应商表现；仓库满足率和发单覆盖率分别反映最终保障与HR发单质量。",
      );
    sheet.cell("A4").value(`${analysis.meta.scope_prefix}月度趋势`);
    sheet.range("A5:G5").value([
      [
        "月份",
        "原始接单兑现率",
        "超发校正率",
        "仓库真实满足率",
        "发单覆盖率",
        "实际发单/真实需求",
        "派遣/真实需求",
      ],
    ]);

    function sumIf(sheetName, sumColumn, criteriaColumn, row, start, end) {
      return `SUMIFS('${sheetName}'!$${sumColumn}$${start}:$${sumColumn}$${end},'${sheetName}'!$${criteriaColumn}$${start}:$${criteriaColumn}$${end},$A${row})`;
    }

    analysis.meta.months.forEach((month, index) => {
      const row = monthStart + index;
      const values = analysis.scope_summary.monthly[index];
      applyRowStyles(sheet, row, styles.month);
      sheet.cell(row, 1).value(month.key);
      const rawDen = sumIf("供应商月度明细", "F", "E", row, ranges.monthly.start, ranges.monthly.end);
      const rawNum = sumIf("供应商月度明细", "H", "E", row, ranges.monthly.start, ranges.monthly.end);
      const adjustedDen = sumIf("供应商月度明细", "J", "E", row, ranges.monthly.start, ranges.monthly.end);
      const adjustedNum = sumIf("供应商月度明细", "K", "E", row, ranges.monthly.start, ranges.monthly.end);
      const demand = sumIf("仓库单元底表", "J", "E", row, ranges.unit.start, ranges.unit.end);
      const warehouseEffective = sumIf("仓库单元底表", "N", "E", row, ranges.unit.start, ranges.unit.end);
      const issueCovered = sumIf("仓库单元底表", "M", "E", row, ranges.unit.start, ranges.unit.end);
      const issued = sumIf("仓库单元底表", "K", "E", row, ranges.unit.start, ranges.unit.end);
      const dispatched = sumIf("仓库单元底表", "L", "E", row, ranges.unit.start, ranges.unit.end);
      [
        [2, rawNum, rawDen, values.raw_rate],
        [3, adjustedNum, adjustedDen, values.adjusted_rate],
        [4, warehouseEffective, demand, values.warehouse_rate],
        [5, issueCovered, demand, values.issue_coverage_rate],
        [6, issued, demand, values.issue_to_demand],
        [7, dispatched, demand, values.dispatch_to_demand],
      ].forEach(([column, numerator, denominator, cached]) => {
        setFormula(
          sheet.cell(row, column),
          ratioFormula(numerator, denominator),
          cached,
        );
      });
    });

    applyRowStyles(sheet, averageRow, styles.average);
    applyRowStyles(sheet, weightedRow, styles.weighted);
    sheet.cell(averageRow, 1).value(`${analysis.meta.month_count}个月算术平均`);
    sheet.cell(weightedRow, 1).value(`${analysis.meta.period_short}累计加权`);
    [
      "raw_rate",
      "adjusted_rate",
      "warehouse_rate",
      "issue_coverage_rate",
      "issue_to_demand",
      "dispatch_to_demand",
    ].forEach((metric, index) => {
      const letter = columnName(index + 2);
      setFormula(
        sheet.cell(averageRow, index + 2),
        `IF(COUNT(${letter}${monthStart}:${letter}${monthEnd})=0,"",AVERAGE(${letter}${monthStart}:${letter}${monthEnd}))`,
        analysis.scope_summary.averages[metric],
      );
    });
    const weighted = analysis.scope_summary.weighted;
    const weightedFormulas = [
      ["供应商月度明细", "H", "F", weighted.raw_rate],
      ["供应商月度明细", "K", "J", weighted.adjusted_rate],
      ["仓库单元底表", "N", "J", weighted.warehouse_rate],
      ["仓库单元底表", "M", "J", weighted.issue_coverage_rate],
      ["仓库单元底表", "K", "J", weighted.issue_to_demand],
      ["仓库单元底表", "L", "J", weighted.dispatch_to_demand],
    ];
    weightedFormulas.forEach(([sheetName, numeratorColumn, denominatorColumn, cached], index) => {
      const sourceRange = sheetName === "供应商月度明细" ? ranges.monthly : ranges.unit;
      const numerator = `SUM('${sheetName}'!$${numeratorColumn}$${sourceRange.start}:$${numeratorColumn}$${sourceRange.end})`;
      const denominator = `SUM('${sheetName}'!$${denominatorColumn}$${sourceRange.start}:$${denominatorColumn}$${sourceRange.end})`;
      setFormula(
        sheet.cell(weightedRow, index + 2),
        ratioFormula(numerator, denominator),
        cached,
      );
    });

    sheet.range(keyTitleRow, 1, keyTitleRow, 8).merged(true);
    applyRowStyles(sheet, keyTitleRow, styles.section);
    sheet.cell(keyTitleRow, 1).value(`${analysis.meta.period_short}关键人日与风险`);
    applyRowStyles(sheet, keyPrimaryRow, styles.keyPrimary);
    applyRowStyles(sheet, keyRiskRow, styles.keyRisk);
    applyRowStyles(sheet, keyRatesRow, styles.keyRates);
    sheet.range(keyPrimaryRow, 1, keyRatesRow, 8).value([
      ["真实需求", null, "实际发单", null, "实际派遣", null, "合理责任", null],
      ["少发人日", null, "超发人日", null, "零需求发单", null, "冗余派遣", null],
      [
        `原始${analysis.meta.period_short}加权`,
        null,
        `校正${analysis.meta.period_short}加权`,
        null,
        `仓库${analysis.meta.period_short}满足`,
        null,
        "内置复核",
        analysis.quality.status,
      ],
    ]);
    setFormula(
      sheet.cell(keyPrimaryRow, 2),
      `SUM('仓库单元底表'!$J$${ranges.unit.start}:$J$${ranges.unit.end})`,
      analysis.clean_totals.true_demand,
    );
    setFormula(
      sheet.cell(keyPrimaryRow, 4),
      `SUM('仓库单元底表'!$K$${ranges.unit.start}:$K$${ranges.unit.end})`,
      analysis.clean_totals.issued,
    );
    setFormula(
      sheet.cell(keyPrimaryRow, 6),
      `SUM('仓库单元底表'!$L$${ranges.unit.start}:$L$${ranges.unit.end})`,
      analysis.clean_totals.dispatched,
    );
    setFormula(
      sheet.cell(keyPrimaryRow, 8),
      `SUM('供应商月度明细'!$J$${ranges.monthly.start}:$J$${ranges.monthly.end})`,
      analysis.clean_totals.reasonable_responsibility,
    );
    setFormula(
      sheet.cell(keyRiskRow, 2),
      `SUM('仓库单元底表'!$R$${ranges.unit.start}:$R$${ranges.unit.end})`,
      analysis.unit_atomic.reduce((sum, row) => sum + number(row.under_issue), 0),
    );
    setFormula(
      sheet.cell(keyRiskRow, 4),
      `SUM('仓库单元底表'!$P$${ranges.unit.start}:$P$${ranges.unit.end})`,
      analysis.quality.over_issue.person_days,
    );
    setFormula(
      sheet.cell(keyRiskRow, 6),
      `SUM('仓库单元底表'!$Q$${ranges.unit.start}:$Q$${ranges.unit.end})`,
      analysis.quality.zero_demand_with_issue.issued_person_days,
    );
    setFormula(
      sheet.cell(keyRiskRow, 8),
      `SUM('仓库单元底表'!$S$${ranges.unit.start}:$S$${ranges.unit.end})`,
      analysis.quality.over_dispatch.person_days,
    );
    setFormula(sheet.cell(keyRatesRow, 2), `B${weightedRow}`, weighted.raw_rate);
    setFormula(sheet.cell(keyRatesRow, 4), `C${weightedRow}`, weighted.adjusted_rate);
    setFormula(sheet.cell(keyRatesRow, 6), `D${weightedRow}`, weighted.warehouse_rate);

    sheet.range(regionTitleRow, 1, regionTitleRow, 8).merged(true);
    applyRowStyles(sheet, regionTitleRow, styles.regionSection);
    sheet
      .cell(regionTitleRow, 1)
      .value(analysis.meta.markets?.length > 1 ? "大区与区域对比" : "区域对比");
    applyRowStyles(sheet, regionHeaderRow, styles.regionHeader);
    if (rowHeights.regionHeader) sheet.row(regionHeaderRow).height(rowHeights.regionHeader);
    sheet.range(regionHeaderRow, 1, regionHeaderRow, 8).value([
      [
        analysis.meta.markets?.length > 1 ? "大区/区域" : "区域",
        "原始月均",
        "校正月均",
        "月均变化",
        "仓库月均",
        "覆盖月均",
        `${analysis.meta.period_short}超发人日`,
        `${analysis.meta.period_short}少发人日`,
      ],
    ]);
    const displayRows = new Map();
    analysis.display_metrics.forEach((row, index) => {
      if (["region_total", "market_total"].includes(row.row_type)) {
        displayRows.set(`${row.row_type}:${row.region}`, 5 + index);
      }
    });
    const warehouseRows = new Map();
    analysis.warehouse_display.forEach((row, index) => {
      if (!["region_total", "market_total"].includes(row.row_type)) return;
      const key = `${row.row_type}:${row.region}`;
      if (!warehouseRows.has(key)) warehouseRows.set(key, []);
      warehouseRows.get(key).push(5 + index);
    });
    const rawAverageColumn = columnName(ranges.rawMain.averageColumn);
    const adjustedAverageColumn = columnName(ranges.adjustedMain.averageColumn);
    comparisonRows.forEach((region, index) => {
      const excelRow = regionStartRow + index;
      applyRowStyles(sheet, excelRow, styles.regionBody);
      styles.regionRightBlank.forEach((styleId, styleIndex) => {
        sheet.cell(excelRow, styleIndex + 9)._styleId = styleId;
      });
      sheet.cell(excelRow, 1).value(region.label || region.region);
      const summaryKey = `${region.row_type || "region_total"}:${region.region}`;
      const displayRow = displayRows.get(summaryKey);
      const warehouseRefs = (warehouseRows.get(summaryKey) || [])
        .map((row) => `'仓库满足与发单质量'!O${row}`)
        .join(",");
      const coverageRefs = (warehouseRows.get(summaryKey) || [])
        .map((row) => `'仓库满足与发单质量'!N${row}`)
        .join(",");
      setFormula(
        sheet.cell(excelRow, 2),
        `'主指标_原始兑现率'!${rawAverageColumn}${displayRow}`,
        region.raw_average,
      );
      setFormula(
        sheet.cell(excelRow, 3),
        `'对照_超发校正率'!${adjustedAverageColumn}${displayRow}`,
        region.adjusted_average,
      );
      setFormula(sheet.cell(excelRow, 4), `C${excelRow}-B${excelRow}`, region.average_lift);
      setFormula(sheet.cell(excelRow, 5), `AVERAGE(${warehouseRefs})`, region.warehouse_average);
      setFormula(sheet.cell(excelRow, 6), `AVERAGE(${coverageRefs})`, region.coverage_average);
      const memberRegions = region.member_regions || [region.region];
      const regionSum = (column) => {
        if (memberRegions.length === 1 && memberRegions[0] === region.region) {
          return `SUMIFS('仓库单元底表'!$${column}$${ranges.unit.start}:$${column}$${ranges.unit.end},'仓库单元底表'!$A$${ranges.unit.start}:$A$${ranges.unit.end},$A${excelRow})`;
        }
        return memberRegions
          .map(
            (member) =>
              `SUMIFS('仓库单元底表'!$${column}$${ranges.unit.start}:$${column}$${ranges.unit.end},'仓库单元底表'!$A$${ranges.unit.start}:$A$${ranges.unit.end},${excelText(member)})`,
          )
          .join("+");
      };
      setFormula(sheet.cell(excelRow, 7), regionSum("P"), region.over_issue);
      setFormula(sheet.cell(excelRow, 8), regionSum("R"), region.under_issue);
    });

    sheet.range(conclusionRow, 1, conclusionRow + 1, 12).merged(true);
    applyRowStyles(sheet, conclusionRow, styles.conclusion);
    applyRowStyles(sheet, conclusionRow + 1, styles.conclusionSecond);
    if (rowHeights.conclusion) sheet.row(conclusionRow).height(rowHeights.conclusion);
    if (rowHeights.conclusionSecond) {
      sheet.row(conclusionRow + 1).height(rowHeights.conclusionSecond);
    }
    sheet
      .cell(conclusionRow, 1)
      .value(
        `重要审查结论：超发校正率并非机械加分。原子单元内校正不会降低履约率，但跨单元汇总后，高履约且严重超发的单元权重会下降，因此有${analysis.quality.adjusted_below_raw_count}个供应商×工种×月份的校正率低于原始率。这反映责任重新加权，不是计算错误。`,
      );
    trimSheet(sheet, conclusionRow + 1);
    return { monthStart, monthEnd, averageRow, weightedRow };
  }

  function setReconciliation(workbook, analysis) {
    const sheet = workbook.sheet("三表对账");
    const headerStyles = captureRowStyles(sheet, 4, 14);
    const bodyStyles = captureRowStyles(sheet, 5, 14);
    const sectionStyles = captureRowStyles(sheet, 24, 14);
    const mismatchHeaderStyles = captureRowStyles(sheet, 25, 14);
    const mismatchBodyStyles = captureRowStyles(sheet, 26, 14);
    if (sheet.range("A24:N24").merged()) sheet.range("A24:N24").merged(false);
    clearValues(sheet);
    sheet.cell("A1").value("基本信息、需求详情、发单详情交叉对账");
    sheet
      .cell("A2")
      .value(
        `区间日期按日展开后，月×区域层面三表需求、发单、派遣应完全一致。组日级源系统展示错位或待发字段异常列于下方；当前关键对账状态为${analysis.quality.status}。`,
      );
    const headers = [
      "月份",
      "区域",
      "基本需求",
      "需求详情需求",
      "需求差",
      "基本发单",
      "需求详情发单",
      "发单详情发单",
      "发单最大差",
      "基本派遣",
      "需求详情派遣",
      "发单详情派遣",
      "派遣最大差",
      "状态",
    ];
    sheet.range("A4:N4").value([headers]);
    applyRowStyles(sheet, 4, headerStyles);
    const start = 5;
    const rows = analysis.quality.month_region_reconciliation;
    const values = rows.map((row) => [
      row.month,
      row.region,
      number(row.basic_demand),
      number(row.demand_detail_demand),
      null,
      number(row.basic_issued),
      number(row.demand_detail_issued),
      number(row.dispatch_detail_issued),
      null,
      number(row.basic_dispatched),
      number(row.demand_detail_dispatched),
      number(row.dispatch_detail_dispatched),
      null,
      null,
    ]);
    const end = fillRows(sheet, start, values, () => bodyStyles);
    rows.forEach((row, index) => {
      const excelRow = start + index;
      setFormula(
        sheet.cell(excelRow, 5),
        `C${excelRow}-D${excelRow}`,
        row.demand_difference,
      );
      setFormula(
        sheet.cell(excelRow, 9),
        `MAX(F${excelRow}:H${excelRow})-MIN(F${excelRow}:H${excelRow})`,
        row.issue_max_difference,
      );
      setFormula(
        sheet.cell(excelRow, 13),
        `MAX(J${excelRow}:L${excelRow})-MIN(J${excelRow}:L${excelRow})`,
        row.dispatch_max_difference,
      );
      setFormula(
        sheet.cell(excelRow, 14),
        `IF(MAX(ABS(E${excelRow}),I${excelRow},M${excelRow})=0,"一致","不一致")`,
        row.status,
      );
    });

    const sectionRow = end + 3;
    sheet.range(sectionRow, 1, sectionRow, 14).merged(true);
    sheet.cell(sectionRow, 1).value("组日级源系统差异");
    applyRowStyles(sheet, sectionRow, sectionStyles);
    const mismatchHeaderRow = sectionRow + 1;
    const mismatchHeaders = [
      "区域",
      "仓库",
      "组",
      "日期",
      "基本需求",
      "详情需求",
      "基本待发",
      "详情待发",
      "基本发单",
      "详情发单",
      "基本待派",
      "详情待派",
      "基本派遣",
      "详情派遣",
    ];
    sheet.range(mismatchHeaderRow, 1, mismatchHeaderRow, 14).value([mismatchHeaders]);
    applyRowStyles(sheet, mismatchHeaderRow, mismatchHeaderStyles);
    const mismatchRows = analysis.quality.basic_vs_demand_mismatches;
    const mismatchStart = mismatchHeaderRow + 1;
    const mismatchValues = mismatchRows.length
      ? mismatchRows.map((row) => [
          row.region,
          row.warehouse,
          row.group,
          row.date,
          number(row.basic_demand),
          number(row.demand_detail_demand),
          number(row.basic_pending_issue),
          number(row.demand_detail_pending_issue),
          number(row.basic_issued),
          number(row.demand_detail_issued),
          number(row.basic_pending_dispatch),
          number(row.demand_detail_pending_dispatch),
          number(row.basic_dispatched),
          number(row.demand_detail_dispatched),
        ])
      : [["无差异", null, null, null, null, null, null, null, null, null, null, null, null, null]];
    const finalRow = fillRows(sheet, mismatchStart, mismatchValues, () => mismatchBodyStyles);
    trimSheet(sheet, finalRow);
    return { start, end, headers };
  }

  function setMethodology(workbook, analysis) {
    const sheet = workbook.sheet("口径与质量审查");
    clearValues(sheet);
    sheet.cell("A1").value("统计口径、第一性原理与对抗性审查");
    sheet
      .cell("A2")
      .value(
        "本页区分可观察事实、责任分摊模型和治理用途，避免把仓库保障、HR发单质量、供应商履约混为一个指标。",
      );
    sheet.cell("A4").value("一、四类指标各回答什么问题");
    sheet.range("A5:D8").value([
      ["指标", "定位", "公式", "主要用途"],
      [
        "原始接单兑现率",
        "主指标",
        analysis.method.raw_rate,
        "供应商对实际收到订单的兑现能力；保留HR超发造成的现实影响",
      ],
      [
        "超发校正履约率",
        "对照",
        analysis.method.adjusted_rate,
        "仅在总发单>真实需求时按发单占比缩减责任，用于判断超发是否稀释表现",
      ],
      [
        "仓库真实满足率 / 发单覆盖率",
        "治理背景",
        `${analysis.method.warehouse_rate}；${analysis.method.issue_coverage}`,
        "分别评价仓库最终保障结果与HR是否把真实需求有效发出",
      ],
    ]);
    sheet.cell("A10").value("二、人日与小数解释");
    sheet
      .cell("A11")
      .value(
        "1人工作1天=1人日；10人工作1天=10人日；同1个人连续工作10天=10人日，但唯一人数仍是1人。区间记录“起始日~结束日，需求10”表示每天10人。原始需求、发单、派遣应保持整数人日；按发单比例折算出的7.92属于“责任人日等值”，不是7.92个人。",
      );
    sheet.cell("A15").value("三、清洗与计算规则");
    const supplement = analysis.clean_totals.valid_supplement;
    const rules = [
      ["最细颗粒度", analysis.method.atomic_grain],
      [
        "日期区间",
        `区间字段按“每天N人”逐日展开，并按文件期间及${analysis.meta.period_end}截断`,
      ],
      [
        "补员",
        `有效补员${supplement}人日并入匹配工种；培训测试单元内补员随整单元剔除`,
      ],
      ["测试单元", analysis.method.test_unit],
      ["零需求发单", analysis.method.zero_demand_issue],
      ["月度平均", analysis.method.average],
      [`${analysis.meta.period_short}累计加权`, analysis.method.period_weighted],
      ["无数据月份", analysis.meta.no_data_rule],
    ];
    sheet.range("A16:B23").value(rules);
    sheet.cell("A25").value("四、数据质量与对抗性审查");
    const criticalDifference = analysis.quality.critical_checks.reduce(
      (sum, check) => sum + check.difference_count,
      0,
    );
    const qualityRows = [
      [
        "原始三表月×区域对账",
        `${analysis.quality.month_region_reconciliation.length}个单元，差异${analysis.quality.month_region_reconciliation_mismatch_count}`,
      ],
      [
        "需求详情 vs 发单详情",
        `最细单元发单/派遣差异${analysis.quality.demand_vs_dispatch_mismatch_count}`,
      ],
      [
        "内置闭合复核",
        `${analysis.quality.critical_checks.length}项关键检查，差异${criticalDifference}，状态${analysis.quality.status}`,
      ],
      [
        "责任闭合",
        `Σ供应商合理责任=min(单元总发单,真实需求)，差异${analysis.quality.responsibility_closure_mismatch_count}`,
      ],
      [
        "测试剔除",
        `完整剔除${analysis.quality.test_unit_count}个培训测试原子单元`,
      ],
      [
        "超发",
        `${analysis.quality.over_issue.unit_count}个正需求单元，超发${analysis.quality.over_issue.person_days}人日`,
      ],
      [
        "零需求发单",
        `${analysis.quality.zero_demand_with_issue.unit_count}个单元，发单${analysis.quality.zero_demand_with_issue.issued_person_days}、派遣${analysis.quality.zero_demand_with_issue.dispatched_person_days}人日`,
      ],
      [
        "未发单缺口",
        `${analysis.quality.positive_demand_no_issue.unit_count}个单元，需求${analysis.quality.positive_demand_no_issue.demand_person_days}人日`,
      ],
      [
        "超派遣",
        `${analysis.quality.over_dispatch.unit_count}个单元，冗余派遣${analysis.quality.over_dispatch.person_days}人日`,
      ],
      [
        "反直觉结果",
        `${analysis.quality.adjusted_below_raw_count}个供应商×工种×月的校正率低于原始率，原因是超发高履约单元的权重被下调，不是计算错误`,
      ],
    ];
    sheet.range("A26:B35").value(qualityRows);
    sheet.cell("A37").value("五、必须保留的模型边界");
    sheet
      .cell("A38")
      .value(
        "比例分摊只能中性地分配“超发责任”，不能证明HR具体多发给了哪一家供应商。若未来系统能够提供预分配配额、供应商确认接单量、发单优先级或时间戳，应优先用该责任依据替代比例模型。当前校正率只能作为对照诊断，不应替代已确认的原始接单兑现率主指标。",
      );
    trimSheet(sheet, 40);
  }

  function setSourceFiles(workbook, analysis) {
    const sheet = workbook.sheet("源文件清单");
    const headerStyles = captureRowStyles(sheet, 4, 6);
    const bodyStyles = captureRowStyles(sheet, 5, 6);
    const sectionStyles = captureRowStyles(sheet, 14, 10);
    const totalHeaderStyles = captureRowStyles(sheet, 15, 10);
    const totalBodyStyles = captureRowStyles(sheet, 16, 10);
    if (sheet.range("A14:J14").merged()) sheet.range("A14:J14").merged(false);
    clearValues(sheet);
    sheet.cell("A1").value("源文件版本锁定与原始总量");
    sheet
      .cell("A2")
      .value(
        "SHA-256用于确认后续复算使用的是同一批原始文件。原始总量含培训测试单元；正式口径在管理摘要和口径页说明。",
      );
    const headers = ["文件", "期间开始", "期间结束", "大小(Bytes)", "修改时间", "SHA-256"];
    sheet.range("A4:F4").value([headers]);
    applyRowStyles(sheet, 4, headerStyles);
    const start = 5;
    const values = analysis.source_files.map((row) => [
      row.source_file,
      row.period_start,
      row.period_end,
      number(row.size_bytes),
      row.modified_at,
      row.sha256,
    ]);
    const end = fillRows(sheet, start, values, () => bodyStyles);
    const sectionRow = end + 3;
    sheet.range(sectionRow, 1, sectionRow, 10).merged(true);
    sheet.cell(sectionRow, 1).value("原始三表月度总量（含测试）");
    applyRowStyles(sheet, sectionRow, sectionStyles);
    const totalHeaderRow = sectionRow + 1;
    const totalHeaders = [
      "月份",
      "基本需求",
      "基本补员",
      "基本发单",
      "基本派遣",
      "需求详情需求",
      "需求详情发单",
      "需求详情派遣",
      "发单详情发单",
      "发单详情派遣",
    ];
    sheet.range(totalHeaderRow, 1, totalHeaderRow, 10).value([totalHeaders]);
    applyRowStyles(sheet, totalHeaderRow, totalHeaderStyles);
    const totalStart = totalHeaderRow + 1;
    const totalValues = analysis.source_totals.map((row) => [
      row.month,
      number(row.basic_demand),
      number(row.basic_supplement),
      number(row.basic_issued),
      number(row.basic_dispatched),
      number(row.demand_detail_demand),
      number(row.demand_detail_issued),
      number(row.demand_detail_dispatched),
      number(row.dispatch_detail_issued),
      number(row.dispatch_detail_dispatched),
    ]);
    const finalRow = fillRows(sheet, totalStart, totalValues, () => totalBodyStyles);
    trimSheet(sheet, finalRow);
    return { start, end, headers };
  }

  async function updateTableXml(workbook, tableName, headers, endRow) {
    const paths = Object.keys(workbook._zip.files).filter((path) =>
      /^xl\/tables\/table\d+\.xml$/.test(path),
    );
    for (const path of paths) {
      const file = workbook._zip.file(path);
      if (!file) continue;
      let xml = await file.async("string");
      if (!new RegExp(`\\bname="${tableName}"`).test(xml)) continue;
      const endColumn = columnName(headers.length);
      const ref = `A4:${endColumn}${endRow}`;
      xml = xml
        .replace(/(<table\b[^>]*\bref=")[^"]+/, `$1${ref}`)
        .replace(/(<autoFilter\b[^>]*\bref=")[^"]+/, `$1${ref}`);
      const columns = headers
        .map(
          (header, index) =>
            `<tableColumn id="${index + 1}" name="${xmlEscape(header)}"></tableColumn>`,
        )
        .join("");
      xml = xml.replace(
        /<tableColumns\b[^>]*>[\s\S]*?<\/tableColumns>/,
        `<tableColumns count="${headers.length}">${columns}</tableColumns>`,
      );
      workbook._zip.file(path, xml);
      return;
    }
    throw new Error(`模板中未找到表格 ${tableName}`);
  }

  function chartPoints(values) {
    return values
      .map((value, index) => {
        if (value == null || !Number.isFinite(Number(value))) return "";
        return `<c:pt idx="${index}"><c:v>${Number(value)}</c:v></c:pt>`;
      })
      .join("");
  }

  function monthStringCache(months) {
    return `<c:strCache><c:ptCount val="${months.length}"/>${months
      .map(
        (month, index) =>
          `<c:pt idx="${index}"><c:v>${xmlEscape(month.key)}</c:v></c:pt>`,
      )
      .join("")}</c:strCache>`;
  }

  async function updateChartXml(workbook, analysis) {
    const path = Object.keys(workbook._zip.files).find((candidate) =>
      /(?:^|\/)charts\/chart1\.xml$/.test(candidate),
    );
    if (!path) return;
    const file = workbook._zip.file(path);
    if (!file) return;
    let xml = await file.async("string");
    const title = `${analysis.meta.scope_prefix}${analysis.meta.period_label}关键履约与保障指标`;
    xml = xml.replace(
      /<a:t>[^<]*关键履约与保障指标<\/a:t>/,
      `<a:t>${xmlEscape(title)}</a:t>`,
    );
    const monthEndRow = 5 + analysis.meta.months.length;
    xml = xml.replace(
      /'管理摘要'!\$([A-E])\$6:\$\1\$\d+/g,
      (_match, column) => `'管理摘要'!$${column}$6:$${column}$${monthEndRow}`,
    );
    const categoryCache = monthStringCache(analysis.meta.months);
    xml = xml.replace(
      /(<c:cat><c:strRef><c:f>'管理摘要'!\$A\$6:\$A\$\d+<\/c:f>)<c:strCache>[\s\S]*?<\/c:strCache>/g,
      `$1${categoryCache}`,
    );
    const series = [
      ["B", analysis.scope_summary.monthly.map((row) => row.raw_rate)],
      ["C", analysis.scope_summary.monthly.map((row) => row.adjusted_rate)],
      ["D", analysis.scope_summary.monthly.map((row) => row.warehouse_rate)],
      ["E", analysis.scope_summary.monthly.map((row) => row.issue_coverage_rate)],
    ];
    series.forEach(([column, values]) => {
      const pattern = new RegExp(
        `(<c:val><c:numRef><c:f>'管理摘要'!\\$${column}\\$6:\\$${column}\\$\\d+<\\/c:f>)<c:numCache>[\\s\\S]*?<\\/c:numCache>`,
      );
      const cache = `<c:numCache><c:formatCode>0.0%</c:formatCode><c:ptCount val="${values.length}"/>${chartPoints(
        values,
      )}</c:numCache>`;
      xml = xml.replace(pattern, `$1${cache}`);
    });
    workbook._zip.file(path, xml);
  }

  async function updateManagementPrintXml(workbook) {
    const workbookXml = await workbook._zip.file("xl/workbook.xml").async("string");
    const relationshipId = workbookXml.match(
      /<sheet\b[^>]*name="管理摘要"[^>]*r:id="([^"]+)"[^>]*\/>/,
    )?.[1];
    if (!relationshipId) throw new Error("未找到管理摘要工作表关系");
    const relationships = await workbook._zip
      .file("xl/_rels/workbook.xml.rels")
      .async("string");
    const relationTag = relationships
      .match(/<Relationship\b[^>]*\/>/g)
      ?.find((tag) => tag.includes(`Id="${relationshipId}"`));
    const target = relationTag?.match(/\bTarget="([^"]+)"/)?.[1];
    if (!target) throw new Error("未找到管理摘要工作表路径");
    const path = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
    let xml = await workbook._zip.file(path).async("string");
    if (/<sheetPr\b[^>]*\/>/.test(xml)) {
      xml = xml.replace(
        /<sheetPr\b([^>]*)\/>/,
        '<sheetPr$1><pageSetUpPr fitToPage="1"/></sheetPr>',
      );
    } else if (/<sheetPr\b[^>]*>/.test(xml)) {
      if (!/<pageSetUpPr\b/.test(xml)) {
        xml = xml.replace(/<sheetPr\b([^>]*)>/, '<sheetPr$1><pageSetUpPr fitToPage="1"/>');
      }
    } else {
      xml = xml.replace(
        /(<worksheet\b[^>]*>)/,
        '$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
      );
    }
    const pageSetup =
      '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>';
    if (/<pageSetup\b[^>]*\/>/.test(xml)) {
      xml = xml.replace(/<pageSetup\b[^>]*\/>/, pageSetup);
    } else if (/<pageMargins\b[^>]*\/>/.test(xml)) {
      xml = xml.replace(/(<pageMargins\b[^>]*\/>)/, `$1${pageSetup}`);
    } else {
      xml = xml.replace(/(<drawing\b|<legacyDrawing\b|<\/worksheet>)/, `${pageSetup}$1`);
    }
    workbook._zip.file(path, xml);
  }

  async function loadTemplate(pathValue) {
    if (root.SupplierFulfillmentTemplateBase64) {
      const binary = root.atob(root.SupplierFulfillmentTemplateBase64);
      const bytes = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index += 1) {
        bytes[index] = binary.charCodeAt(index);
      }
      return bytes;
    }
    if (typeof fetch !== "function") {
      throw new Error("当前环境无法加载报表模板");
    }
    const response = await fetch(pathValue || TEMPLATE_PATH);
    if (!response.ok) throw new Error(`报表模板加载失败：HTTP ${response.status}`);
    return new Uint8Array(await response.arrayBuffer());
  }

  async function buildReport(analysis, templateData) {
    if (!analysis?.meta || !analysis?.supplier_atomic) {
      throw new Error("缺少有效的履约率分析结果");
    }
    const source = templateData || (await loadTemplate(TEMPLATE_PATH));
    const workbook = await getXlsxPopulate().fromDataAsync(source);

    const supplierAtomic = setSupplierAtomic(workbook, analysis);
    const unitAtomic = setUnitAtomic(workbook, analysis);
    const monthly = setSupplierMonthly(workbook, analysis, supplierAtomic);
    const rawMain = setMainSheet(workbook, analysis, "raw", monthly);
    const adjustedMain = setMainSheet(workbook, analysis, "adjusted", monthly);
    const warehouse = setWarehouseSummary(workbook, analysis, unitAtomic, monthly);
    setDashboard(workbook, analysis, {
      supplierAtomic,
      unit: unitAtomic,
      monthly,
      rawMain,
      adjustedMain,
      warehouse,
    });
    const reconciliation = setReconciliation(workbook, analysis);
    setMethodology(workbook, analysis);
    const sourceFiles = setSourceFiles(workbook, analysis);

    const tables = [
      ["RawFulfillmentScorecard", rawMain],
      ["AdjustedFulfillmentScorecard", adjustedMain],
      ["SupplierMonthlyMetrics", monthly],
      ["WarehouseAndAllocationSummary", warehouse],
      ["SupplierAtomicResponsibility", supplierAtomic],
      ["WarehouseAtomicFacts", unitAtomic],
      ["MonthlyThreeWayReconciliation", reconciliation],
      ["SourceFileRegistry", sourceFiles],
    ];
    for (const [name, range] of tables) {
      await updateTableXml(workbook, name, range.headers, range.end);
    }
    await updateChartXml(workbook, analysis);

    await workbook.outputAsync();
    await updateManagementPrintXml(workbook);
    const supportsBlob = typeof Blob !== "undefined";
    let blob = await workbook._zip.generateAsync({
      type: supportsBlob ? "blob" : "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
      ...(supportsBlob
        ? { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
        : {}),
    });
    if (!supportsBlob && typeof Buffer !== "undefined" && !Buffer.isBuffer(blob)) {
      blob = Buffer.from(blob);
    }
    return { blob, fileName: analysis.meta.export_filename };
  }

  function download(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    root.setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  return {
    buildReport,
    download,
    loadTemplate,
    columnName,
    setFormula,
    TABLE_NAMES,
  };
});

(function (root, factory) {
  const api = factory(root);
  if (typeof module === "object" && module.exports) module.exports = api;
  root.RetentionReportExporter = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function (root) {
  "use strict";

  const COUNTRY_ORDER = ["德国", "意大利", "捷克", "法国", "波兰", "英国", "西班牙"];
  const GLOBAL_REGION_ORDER = ["欧洲", "美洲", "亚太", "未识别"];
  const TABLE_CONFIG = {
    "国家汇总": { path: "xl/tables/table1.xml", endColumn: "O" },
    "供应商留存率": { path: "xl/tables/table2.xml" },
    "人员明细": { path: "xl/tables/table3.xml", endColumn: "AP" },
    "异常复核": { path: "xl/tables/table4.xml", endColumn: "O" },
  };

  function getXlsxPopulate() {
    if (root.XlsxPopulate) return root.XlsxPopulate;
    if (typeof require === "function") return require("./vendor/xlsx-populate-no-encryption.min.js");
    throw new Error("报表模板引擎未加载");
  }

  function decodeBase64(base64) {
    if (typeof Buffer !== "undefined") return new Uint8Array(Buffer.from(base64, "base64"));
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  }

  function columnName(number) {
    let result = "";
    let value = number;
    while (value > 0) {
      value -= 1;
      result = String.fromCharCode(65 + (value % 26)) + result;
      value = Math.floor(value / 26);
    }
    return result;
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
    const parts = String(iso).slice(0, 10).split("-").map(Number);
    if (parts.length !== 3 || parts.some((part) => !Number.isFinite(part))) return null;
    return new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0, 0);
  }

  function dateFromDay(day, dayToISO) {
    return day == null ? null : localDate(dayToISO(day));
  }

  function excelSerial(day) {
    return day == null ? null : day + 25569.5;
  }

  function setFormula(cell, formula, cachedValue) {
    const normalizedValue = typeof cachedValue === "number" && Number.isFinite(cachedValue)
      ? Number(cachedValue.toPrecision(15))
      : cachedValue;
    cell.formula(formula);
    const attributes = { ...(cell._remainingAttributes || {}) };
    delete attributes.t;
    if (typeof normalizedValue === "string") attributes.t = "str";
    cell._remainingAttributes = attributes;
    cell._remainingChildren = [{ name: "v", children: [normalizedValue == null ? "" : normalizedValue] }];
    return cell;
  }

  function fillRange(sheet, range, value, rows, columns) {
    sheet.range(range).value(Array.from({ length: rows }, () => Array(columns).fill(value)));
  }

  function chineseDate(iso) {
    if (!iso) return "";
    const [year, month, day] = String(iso).slice(0, 10).split("-").map(Number);
    return `${year}年${month}月${day}日`;
  }

  function monthSpan(startMonth, endMonth) {
    return startMonth === endMonth ? `${startMonth}月` : `${startMonth}-${endMonth}月`;
  }

  function attendanceProfiles(context) {
    return (context.profiles.attendance || []).slice().sort((left, right) => {
      return String(left.dateMin || "").localeCompare(String(right.dateMin || ""));
    });
  }

  function coverageMonthSpan(context) {
    if (!context.attendanceStore || context.attendanceStore.coverageMin == null || context.attendanceStore.coverageMax == null) return "全部考勤源";
    const start = context.dayToISO(context.attendanceStore.coverageMin);
    const end = context.dayToISO(context.attendanceStore.coverageMax);
    const startYear = Number(start.slice(0, 4));
    const endYear = Number(end.slice(0, 4));
    const startMonth = Number(start.slice(5, 7));
    const endMonth = Number(end.slice(5, 7));
    return startYear === endYear ? monthSpan(startMonth, endMonth) : `${start}至${end}`;
  }

  function reportQualityFlags(row, context) {
    const flags = row.qualityFlags.slice();
    if (row.manualFlags.includes("T窗口出现多个供应商") && !flags.includes("T窗口出现多个供应商")) {
      flags.push("T窗口出现多个供应商");
    }
    return flags
      .flatMap((flag) => {
        if (flag === "全部考勤源无有效出勤") return [`${coverageMonthSpan(context)}无有效出勤`];
        if (flag !== "日期恰好等于T") return [flag];
        const boundaryFlags = [];
        if (row.latestDispatch === row.observationDay) boundaryFlags.push("最晚派遣日期恰好等于T");
        if (row.latestEnd === row.observationDay) boundaryFlags.push("最近结束日期恰好等于T");
        return boundaryFlags.length ? boundaryFlags : [flag];
      })
      .join("；");
  }

  function countryRank(country) {
    const index = COUNTRY_ORDER.indexOf(country);
    return index >= 0 ? index : COUNTRY_ORDER.length;
  }

  function orderCountries(rows) {
    return rows.slice().sort((left, right) => {
      const rank = countryRank(left.country) - countryRank(right.country);
      if (rank) return rank;
      return String(left.country).localeCompare(String(right.country), "zh-CN");
    });
  }

  function orderSuppliers(rows) {
    return rows.slice().sort((left, right) => {
      const country = countryRank(left.country) - countryRank(right.country);
      if (country) return country;
      if ((right.retentionRate ?? -1) !== (left.retentionRate ?? -1)) {
        return (right.retentionRate ?? -1) - (left.retentionRate ?? -1);
      }
      if (right.denominator !== left.denominator) return right.denominator - left.denominator;
      return String(left.supplierName).localeCompare(String(right.supplierName), "zh-CN");
    });
  }

  function globalRegionRank(value) {
    const index = GLOBAL_REGION_ORDER.indexOf(value);
    return index >= 0 ? index : GLOBAL_REGION_ORDER.length;
  }

  function orderAreas(rows) {
    return rows.slice().sort((left, right) => {
      const globalRegion = globalRegionRank(left.globalRegion) - globalRegionRank(right.globalRegion);
      if (globalRegion) return globalRegion;
      const country = String(left.country).localeCompare(String(right.country), "zh-CN");
      if (country) return country;
      return String(left.region).localeCompare(String(right.region), "zh-CN");
    });
  }

  function orderGlobalSuppliers(rows) {
    return rows.slice().sort((left, right) => {
      const globalRegion = globalRegionRank(left.globalRegion) - globalRegionRank(right.globalRegion);
      if (globalRegion) return globalRegion;
      const country = String(left.country).localeCompare(String(right.country), "zh-CN");
      if (country) return country;
      const region = String(left.region).localeCompare(String(right.region), "zh-CN");
      if (region) return region;
      if ((right.retentionRate ?? -1) !== (left.retentionRate ?? -1)) {
        return (right.retentionRate ?? -1) - (left.retentionRate ?? -1);
      }
      if (right.denominator !== left.denominator) return right.denominator - left.denominator;
      return String(left.supplierName).localeCompare(String(right.supplierName), "zh-CN");
    });
  }

  function monthLabel(month) {
    return `${Number(String(month).slice(5, 7))}月`;
  }

  function cohortTitle(startLabel, endLabel) {
    const start = startLabel.split("-").map(Number);
    const end = endLabel.split("-").map(Number);
    if (start[0] === end[0]) {
      if (start[1] === end[1]) return `${start[0]}年${start[1]}月`;
      return `${start[0]}年${start[1]}-${end[1]}月`;
    }
    return `${startLabel}至${endLabel}`;
  }

  function reportTag(startLabel, endLabel) {
    const start = startLabel.split("-").map(Number);
    const end = endLabel.split("-").map(Number);
    if (start[0] === end[0] && [1, 4, 7, 10].includes(start[1])) {
      const quarter = Math.floor((start[1] - 1) / 3) + 1;
      const quarterEndMonth = quarter * 3;
      const quarterEndDay = new Date(end[0], quarterEndMonth, 0).getDate();
      if (start[2] === 1 && end[1] === quarterEndMonth && end[2] === quarterEndDay) return `${start[0]}Q${quarter}`;
    }
    return `${startLabel.replace(/-/g, "")}-${endLabel.replace(/-/g, "")}`;
  }

  function suggestedAction(row) {
    if (row.manualFlags.includes("系统派遣覆盖T但窗口无有效出勤")) return "核实批准休假、未排班或系统漏结束；无证据则维持未留存";
    if (row.manualFlags.includes("T边界日期需确认")) return "确认结束日期是否为当日结束及公司边界口径";
    return "核对派遣记录、结束记录与考勤原始凭证";
  }

  function templateSourceColumn(sheetName, column, totalColumns) {
    if (sheetName !== "供应商留存率" || totalColumns <= 28) return column;
    if (column <= 18) return column;
    if (column === totalColumns) return 28;
    return 19 + ((column - 19) % 3);
  }

  function fillTableSheet(workbook, sheetName, rows, sourceColumnResolver) {
    const sheet = workbook.sheet(sheetName);
    const rowCount = rows.length;
    const columnCount = rows[0].length;
    const oldRows = sheet._rows.length - 1;
    const oldColumns = sheet.usedRange() ? sheet.usedRange().endCell().columnNumber() : columnCount;
    const headerStyles = [];
    const bodyStyles = [];
    const widths = [];
    for (let column = 1; column <= columnCount; column += 1) {
      const requestedSource = sourceColumnResolver
        ? sourceColumnResolver(column, columnCount, oldColumns)
        : templateSourceColumn(sheetName, column, columnCount);
      const source = Math.min(requestedSource, oldColumns);
      headerStyles[column] = sheet.cell(1, source)._styleId;
      bodyStyles[column] = sheet.cell(2, source)._styleId;
      widths[column] = sheet.column(source).width();
    }

    if (oldRows > 0) sheet.range(1, 1, oldRows, Math.max(oldColumns, columnCount)).value(null);
    const normalizedRows = rows.map((row) => row.map((value) => value === "" ? null : value));
    sheet.range(1, 1, rowCount, columnCount).value(normalizedRows);

    for (let column = 1; column <= columnCount; column += 1) {
      sheet.cell(1, column)._styleId = headerStyles[column];
      if (widths[column]) sheet.column(column).width(widths[column]);
    }
    for (let row = 2; row <= rowCount; row += 1) {
      for (let column = 1; column <= columnCount; column += 1) {
        sheet.cell(row, column)._styleId = bodyStyles[column];
      }
    }

    sheet._rows.length = rowCount + 1;
    sheet._sheetDataNode.children = sheet._rows;
    return sheet;
  }

  async function updateTableXml(workbook, path, headers, rowCount) {
    const file = workbook._zip.file(path);
    if (!file) return;
    let xml = await file.async("string");
    const endColumn = columnName(headers.length);
    xml = xml.replace(/(<table\b[^>]*\bref=")[^"]+/, `$1A1:${endColumn}${rowCount}`);
    const columns = headers
      .map((header, index) => `<tableColumn id="${index + 1}" name="${xmlEscape(header)}"/>`)
      .join("");
    xml = xml.replace(
      /<tableColumns\b[^>]*>[\s\S]*?<\/tableColumns>/,
      `<tableColumns count="${headers.length}">${columns}</tableColumns>`
    );
    workbook._zip.file(path, xml);
  }

  async function updateChartXml(workbook, countries, periodDays, globalReport) {
    const path = "xl/charts/chart1.xml";
    const file = workbook._zip.file(path);
    if (!file) return;
    let xml = await file.async("string");
    const endRow = 24 + countries.length;
    const stringPoints = countries
      .map((row, index) => {
        const label = globalReport ? row.region : row.country;
        return `<c:pt idx="${index}"><c:v>${xmlEscape(label)}</c:v></c:pt>`;
      })
      .join("");
    const numberPoints = countries
      .map((row, index) => `<c:pt idx="${index}"><c:v>${row.retentionRate == null ? 0 : row.retentionRate}</c:v></c:pt>`)
      .join("");
    const chartTitle = globalReport ? `各运营区域${periodDays}天毛留存率` : `各国${periodDays}天毛留存率`;
    const seriesName = `${periodDays}天留存率`;
    const seriesCache = `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${xmlEscape(seriesName)}</c:v></c:pt></c:strCache>`;
    const categoryCache = `<c:strCache><c:ptCount val="${countries.length}"/>${stringPoints}</c:strCache>`;
    const valueCache = `<c:numCache><c:formatCode>0.0%</c:formatCode><c:ptCount val="${countries.length}"/>${numberPoints}</c:numCache>`;
    xml = xml
      .replace(/<a:t>\u5404\u56fd\d+\u5929\u6bdb\u7559\u5b58\u7387<\/a:t>/, `<a:t>${chartTitle}</a:t>`)
      .replace(/<c:f>&quot;\d+\u5929\u7559\u5b58\u7387&quot;<\/c:f>/, `<c:f>&quot;${seriesName}&quot;</c:f>`)
      .replace(/<c:grouping val="none"\/>/g, '<c:grouping val="clustered"/>')
      .replace(/\u7edf\u8ba1\u603b\u89c8!\$J\$25:\$J\$\d+/g, `统计总览!$J$25:$J$${endRow}`)
      .replace(/\u7edf\u8ba1\u603b\u89c8!\$K\$25:\$K\$\d+/g, `统计总览!$K$25:$K$${endRow}`)
      .replace(/(<c:tx><c:strRef>[\s\S]*?)<c:strCache>[\s\S]*?<\/c:strCache>([\s\S]*?<\/c:strRef><\/c:tx>)/, `$1${seriesCache}$2`)
      .replace(/(<c:cat><c:strRef>[\s\S]*?)<c:strCache>[\s\S]*?<\/c:strCache>([\s\S]*?<\/c:strRef><\/c:cat>)/, `$1${categoryCache}$2`)
      .replace(/(<c:val><c:numRef>[\s\S]*?)<c:numCache>[\s\S]*?<\/c:numCache>([\s\S]*?<\/c:numRef><\/c:val>)/, `$1${valueCache}$2`);
    workbook._zip.file(path, xml);
  }

  function setOverview(workbook, context, countries) {
    const { aggregate: summary, periodDays, windowDays, startLabel, endLabel, attendanceStore } = context;
    const sheet = workbook.sheet("统计总览");
    sheet.usedRange().value(null);
    const countryLastRow = countries.length + 1;
    const peopleLastRow = context.results.length + 1;
    const title = `欧洲劳务工供应商${periodDays}天留存率报表（${cohortTitle(startLabel, endLabel)}首次派遣队列）`;
    const subtitle = `毛留存率口径｜T=首次派遣日期+${periodDays}天｜有效出勤窗口=T-${windowDays - 1}日至T｜数据截至${chineseDate(context.dayToISO(attendanceStore.coverageMax))}`;
    fillRange(sheet, "A1:N1", title, 1, 14);
    fillRange(sheet, "A2:N2", subtitle, 1, 14);
    sheet.range("A4:G4").value([["原始队列人数", "计分分母", "留存人数", "未留存人数", `${periodDays}天留存率`, "需人工复核", "粗暴剔除误伤候选"]]);
    setFormula(sheet.cell("A5"), `SUM(国家汇总!$C$2:$C$${countryLastRow})`, summary.cohort);
    setFormula(sheet.cell("B5"), `SUM(国家汇总!$E$2:$E$${countryLastRow})`, summary.denominator);
    setFormula(sheet.cell("C5"), `SUM(国家汇总!$F$2:$F$${countryLastRow})`, summary.retained);
    setFormula(sheet.cell("D5"), `SUM(国家汇总!$G$2:$G$${countryLastRow})`, summary.notRetained);
    setFormula(sheet.cell("E5"), 'IFERROR(C5/B5,"")', summary.retentionRate == null ? "" : summary.retentionRate);
    setFormula(sheet.cell("F5"), `SUM(国家汇总!$N$2:$N$${countryLastRow})`, summary.manualReview);
    setFormula(sheet.cell("G5"), `COUNTIF(人员明细!$V$2:$V$${peopleLastRow},"是")`, summary.blindEliminationRisk);

    sheet.range("A8:E8").value([["国家", "计分分母", "留存人数", "未留存人数", `${periodDays}天留存率`]]);
    countries.forEach((country, index) => {
      const row = 9 + index;
      const source = 2 + index;
      setFormula(sheet.cell(row, 1), `国家汇总!A${source}`, country.country);
      setFormula(sheet.cell(row, 2), `国家汇总!E${source}`, country.denominator);
      setFormula(sheet.cell(row, 3), `国家汇总!F${source}`, country.retained);
      setFormula(sheet.cell(row, 4), `国家汇总!G${source}`, country.notRetained);
      setFormula(sheet.cell(row, 5), `国家汇总!H${source}`, country.retentionRate == null ? "" : country.retentionRate);
    });

    fillRange(sheet, "A18:E18", "核心结论", 1, 5);
    fillRange(sheet, "A19:E22", `本报表按个人T判断，不采用“凡出现在${coverageMonthSpan(context)}最晚派遣表即全部剔除”的粗暴算法。当前结果为毛留存率，尚未按我方减量、批准休假或雇主转换等责任原因豁免；所有系统覆盖T但${windowDays}日无有效出勤、日期边界和字段冲突病例均进入【异常复核】。`, 4, 5);
    sheet.range("J24:K24").value([["国家", `${periodDays}天留存率`]]);
    countries.forEach((country, index) => {
      const row = 25 + index;
      const source = 2 + index;
      setFormula(sheet.cell(row, 10), `国家汇总!A${source}`, country.country);
      setFormula(sheet.cell(row, 11), `国家汇总!H${source}`, country.retentionRate == null ? "" : country.retentionRate);
    });
  }

  function globalShiftOneSource(column) {
    return column <= 2 ? 1 : column - 1;
  }

  function globalSupplierSource(column, totalColumns) {
    if (column === 1) return 1;
    if (column <= 19) return column - 1;
    if (column === totalColumns) return 28;
    return 19 + ((column - 20) % 3);
  }

  function globalReviewSource(column) {
    return column <= 3 ? 1 : column - 2;
  }

  function setGlobalOverview(workbook, context, areas) {
    const { aggregate: summary, periodDays, windowDays, startLabel, endLabel, attendanceStore, scopeLabel } = context;
    const sheet = workbook.sheet("统计总览");
    const headerSourceColumns = [1, 1, 2, 3, 4, 5];
    const tableHeaderStyles = headerSourceColumns.map((column) => sheet.cell(8, column)._styleId);
    const tableBodyStyles = headerSourceColumns.map((column) => sheet.cell(9, column)._styleId);
    const chartHeaderStyles = [sheet.cell(24, 10)._styleId, sheet.cell(24, 11)._styleId];
    const chartBodyStyles = [sheet.cell(25, 10)._styleId, sheet.cell(25, 11)._styleId];
    const coreHeaderStyle = sheet.cell(18, 1)._styleId;
    const coreBodyStyle = sheet.cell(19, 1)._styleId;
    sheet.range("A18:E18").merged(false);
    sheet.range("A19:E22").merged(false);
    sheet.usedRange().value(null);
    const areaLastRow = areas.length + 1;
    const peopleLastRow = context.results.length + 1;
    const title = `${scopeLabel}劳务工供应商${periodDays}天留存率报表（${cohortTitle(startLabel, endLabel)}首次派遣队列）`;
    const subtitle = `毛留存率口径｜T=首次派遣日期+${periodDays}天｜有效出勤窗口=T-${windowDays - 1}日至T｜数据截至${chineseDate(context.dayToISO(attendanceStore.coverageMax))}`;
    fillRange(sheet, "A1:N1", title, 1, 14);
    fillRange(sheet, "A2:N2", subtitle, 1, 14);
    sheet.range("A4:G4").value([["原始队列人数", "计分分母", "留存人数", "未留存人数", `${periodDays}天留存率`, "需人工复核", "粗暴剔除误伤候选"]]);
    setFormula(sheet.cell("A5"), `SUM(区域汇总!$D$2:$D$${areaLastRow})`, summary.cohort);
    setFormula(sheet.cell("B5"), `SUM(区域汇总!$F$2:$F$${areaLastRow})`, summary.denominator);
    setFormula(sheet.cell("C5"), `SUM(区域汇总!$G$2:$G$${areaLastRow})`, summary.retained);
    setFormula(sheet.cell("D5"), `SUM(区域汇总!$H$2:$H$${areaLastRow})`, summary.notRetained);
    setFormula(sheet.cell("E5"), 'IFERROR(C5/B5,"")', summary.retentionRate == null ? "" : summary.retentionRate);
    setFormula(sheet.cell("F5"), `SUM(区域汇总!$O$2:$O$${areaLastRow})`, summary.manualReview);
    setFormula(sheet.cell("G5"), `COUNTIF(人员明细!$W$2:$W$${peopleLastRow},"是")`, summary.blindEliminationRisk);

    sheet.range("A8:F8").value([["大区", "国家/运营区域", "计分分母", "留存人数", "未留存人数", `${periodDays}天留存率`]]);
    for (let column = 1; column <= 6; column += 1) {
      sheet.cell(8, column)._styleId = tableHeaderStyles[column - 1];
    }
    sheet.column(1).width(12);
    sheet.column(2).width(28);
    for (let column = 3; column <= 6; column += 1) sheet.column(column).width(14);
    areas.forEach((area, index) => {
      const row = 9 + index;
      const source = 2 + index;
      for (let column = 1; column <= 6; column += 1) {
        sheet.cell(row, column)._styleId = tableBodyStyles[column - 1];
      }
      setFormula(sheet.cell(row, 1), `区域汇总!A${source}`, area.globalRegion);
      setFormula(sheet.cell(row, 2), `区域汇总!B${source}&"｜"&区域汇总!C${source}`, `${area.country}｜${area.region}`);
      setFormula(sheet.cell(row, 3), `区域汇总!F${source}`, area.denominator);
      setFormula(sheet.cell(row, 4), `区域汇总!G${source}`, area.retained);
      setFormula(sheet.cell(row, 5), `区域汇总!H${source}`, area.notRetained);
      setFormula(sheet.cell(row, 6), `区域汇总!I${source}`, area.retentionRate == null ? "" : area.retentionRate);
    });

    const coreStart = Math.max(18, 10 + areas.length);
    for (let column = 1; column <= 6; column += 1) {
      sheet.cell(coreStart, column)._styleId = coreHeaderStyle;
      for (let row = coreStart + 1; row <= coreStart + 4; row += 1) {
        sheet.cell(row, column)._styleId = coreBodyStyle;
      }
    }
    sheet.range(`A${coreStart}:F${coreStart}`).value(null).merged(true);
    sheet.range(`A${coreStart + 1}:F${coreStart + 4}`).value(null).merged(true);
    sheet.cell(coreStart, 1).value("核心结论");
    sheet.cell(coreStart + 1, 1).value(`本报表按个人T判断，不采用“凡出现在${coverageMonthSpan(context)}最晚派遣表即全部剔除”的粗暴算法。当前结果为毛留存率，尚未按我方减量、批准休假或雇主转换等责任原因豁免；所有系统覆盖T但${windowDays}日无有效出勤、日期边界和字段冲突病例均进入【异常复核】。`);
    sheet.row(coreStart).height(28);
    for (let row = coreStart + 1; row <= coreStart + 4; row += 1) sheet.row(row).height(28);
    sheet.range("J24:K24").value([["运营区域", `${periodDays}天留存率`]]);
    sheet.cell(24, 10)._styleId = chartHeaderStyles[0];
    sheet.cell(24, 11)._styleId = chartHeaderStyles[1];
    areas.forEach((area, index) => {
      const row = 25 + index;
      const source = 2 + index;
      sheet.cell(row, 10)._styleId = chartBodyStyles[0];
      sheet.cell(row, 11)._styleId = chartBodyStyles[1];
      setFormula(sheet.cell(row, 10), `区域汇总!C${source}`, area.region);
      setFormula(sheet.cell(row, 11), `区域汇总!I${source}`, area.retentionRate == null ? "" : area.retentionRate);
    });
  }

  function setGlobalAreaSheet(workbook, context, areas, suppliers) {
    const { results, periodDays, windowDays } = context;
    const peopleLastRow = results.length + 1;
    const supplierLastRow = suppliers.length + 1;
    const strictReference = context.aggregate.immature === 0 && context.aggregate.pending === 0;
    const areaKey = (row) => `${row.globalRegion}\u0001${row.country}\u0001${row.region}`;
    const supplierCounts = new Map();
    suppliers.forEach((row) => supplierCounts.set(areaKey(row), (supplierCounts.get(areaKey(row)) || 0) + 1));
    const header = ["大区", "国家", "运营区域", "原始队列人数", "豁免人数", "计分分母", "留存人数", "未留存人数", `${periodDays}天留存率`, "95%CI下限", "95%CI上限", "L<T人数", "T前结束人数", `${windowDays}日无有效出勤`, "需人工复核", "供应商数"];
    const rows = [header, ...areas.map((row) => [row.globalRegion, row.country, row.region, row.cohort, row.exempt, row.denominator, row.retained, row.notRetained, row.retentionRate, row.ciLow, row.ciHigh, row.latestBeforeT, row.endedBeforeT, row.noValidAttendance, row.manualReview, supplierCounts.get(areaKey(row)) || 0])];
    const sheet = fillTableSheet(workbook, "区域汇总", rows, globalShiftOneSource);
    areas.forEach((area, index) => {
      const row = index + 2;
      const base = `人员明细!$A$2:$A$${peopleLastRow},A${row},人员明细!$B$2:$B$${peopleLastRow},B${row},人员明细!$C$2:$C$${peopleLastRow},C${row}`;
      setFormula(sheet.cell(row, 4), `COUNTIFS(${base})`, area.cohort);
      setFormula(sheet.cell(row, 5), `COUNTIFS(${base},人员明细!$AO$2:$AO$${peopleLastRow},"豁免")`, area.exempt);
      setFormula(sheet.cell(row, 6), strictReference ? `D${row}-E${row}` : `G${row}+H${row}`, area.denominator);
      setFormula(sheet.cell(row, 7), `COUNTIFS(${base},人员明细!$AO$2:$AO$${peopleLastRow},"留存")`, area.retained);
      setFormula(sheet.cell(row, 8), `COUNTIFS(${base},人员明细!$AO$2:$AO$${peopleLastRow},"未留存")`, area.notRetained);
      setFormula(sheet.cell(row, 9), `IFERROR(G${row}/F${row},"")`, area.retentionRate == null ? "" : area.retentionRate);
      setFormula(sheet.cell(row, 10), `IF(F${row}=0,"",MAX(0,(I${row}+3.8416/(2*F${row})-1.96*SQRT(I${row}*(1-I${row})/F${row}+3.8416/(4*F${row}^2)))/(1+3.8416/F${row})))`, area.ciLow == null ? "" : area.ciLow);
      setFormula(sheet.cell(row, 11), `IF(F${row}=0,"",MIN(1,(I${row}+3.8416/(2*F${row})+1.96*SQRT(I${row}*(1-I${row})/F${row}+3.8416/(4*F${row}^2)))/(1+3.8416/F${row})))`, area.ciHigh == null ? "" : area.ciHigh);
      setFormula(sheet.cell(row, 12), `COUNTIFS(${base},人员明细!$Z$2:$Z$${peopleLastRow},"最晚派遣日期早于T")`, area.latestBeforeT);
      setFormula(sheet.cell(row, 13), `COUNTIFS(${base},人员明细!$Z$2:$Z$${peopleLastRow},"T前结束且无结束后有效出勤")`, area.endedBeforeT);
      setFormula(sheet.cell(row, 14), `COUNTIFS(${base},人员明细!$Z$2:$Z$${peopleLastRow},"T前${windowDays}日无有效出勤")`, area.noValidAttendance);
      setFormula(sheet.cell(row, 15), `COUNTIFS(${base},人员明细!$AA$2:$AA$${peopleLastRow},"是")`, area.manualReview);
      setFormula(sheet.cell(row, 16), `COUNTIFS(供应商留存率!$A$2:$A$${supplierLastRow},A${row},供应商留存率!$B$2:$B$${supplierLastRow},B${row},供应商留存率!$C$2:$C$${supplierLastRow},C${row})`, supplierCounts.get(areaKey(area)) || 0);
    });
    return { header, rows };
  }

  function setCountrySheet(workbook, context, countries, suppliers) {
    const { results, periodDays, windowDays } = context;
    const peopleLastRow = results.length + 1;
    const supplierLastRow = suppliers.length + 1;
    const strictReference = context.aggregate.immature === 0 && context.aggregate.pending === 0;
    const supplierCounts = new Map();
    suppliers.forEach((row) => supplierCounts.set(row.country, (supplierCounts.get(row.country) || 0) + 1));
    const header = ["国家", "区域", "原始队列人数", "豁免人数", "计分分母", "留存人数", "未留存人数", `${periodDays}天留存率`, "95%CI下限", "95%CI上限", "L<T人数", "T前结束人数", `${windowDays}日无有效出勤`, "需人工复核", "供应商数"];
    const rows = [header, ...countries.map((row) => [row.country, row.region, row.cohort, row.exempt, row.denominator, row.retained, row.notRetained, row.retentionRate, row.ciLow, row.ciHigh, row.latestBeforeT, row.endedBeforeT, row.noValidAttendance, row.manualReview, supplierCounts.get(row.country) || 0])];
    const sheet = fillTableSheet(workbook, "国家汇总", rows);
    countries.forEach((country, index) => {
      const row = index + 2;
      setFormula(sheet.cell(row, 3), `COUNTIF(人员明细!$A$2:$A$${peopleLastRow},A${row})`, country.cohort);
      setFormula(sheet.cell(row, 4), `COUNTIFS(人员明细!$A$2:$A$${peopleLastRow},A${row},人员明细!$AN$2:$AN$${peopleLastRow},"豁免")`, country.exempt);
      setFormula(sheet.cell(row, 5), strictReference ? `C${row}-D${row}` : `F${row}+G${row}`, country.denominator);
      setFormula(sheet.cell(row, 6), `COUNTIFS(人员明细!$A$2:$A$${peopleLastRow},A${row},人员明细!$AN$2:$AN$${peopleLastRow},"留存")`, country.retained);
      setFormula(sheet.cell(row, 7), `COUNTIFS(人员明细!$A$2:$A$${peopleLastRow},A${row},人员明细!$AN$2:$AN$${peopleLastRow},"未留存")`, country.notRetained);
      setFormula(sheet.cell(row, 8), `IFERROR(F${row}/E${row},"")`, country.retentionRate == null ? "" : country.retentionRate);
      setFormula(sheet.cell(row, 9), `IF(E${row}=0,"",MAX(0,(H${row}+3.8416/(2*E${row})-1.96*SQRT(H${row}*(1-H${row})/E${row}+3.8416/(4*E${row}^2)))/(1+3.8416/E${row})))`, country.ciLow == null ? "" : country.ciLow);
      setFormula(sheet.cell(row, 10), `IF(E${row}=0,"",MIN(1,(H${row}+3.8416/(2*E${row})+1.96*SQRT(H${row}*(1-H${row})/E${row}+3.8416/(4*E${row}^2)))/(1+3.8416/E${row})))`, country.ciHigh == null ? "" : country.ciHigh);
      setFormula(sheet.cell(row, 11), `COUNTIFS(人员明细!$A$2:$A$${peopleLastRow},A${row},人员明细!$Y$2:$Y$${peopleLastRow},"最晚派遣日期早于T")`, country.latestBeforeT);
      setFormula(sheet.cell(row, 12), `COUNTIFS(人员明细!$A$2:$A$${peopleLastRow},A${row},人员明细!$Y$2:$Y$${peopleLastRow},"T前结束且无结束后有效出勤")`, country.endedBeforeT);
      setFormula(sheet.cell(row, 13), `COUNTIFS(人员明细!$A$2:$A$${peopleLastRow},A${row},人员明细!$Y$2:$Y$${peopleLastRow},"T前${windowDays}日无有效出勤")`, country.noValidAttendance);
      setFormula(sheet.cell(row, 14), `COUNTIFS(人员明细!$A$2:$A$${peopleLastRow},A${row},人员明细!$Z$2:$Z$${peopleLastRow},"是")`, country.manualReview);
      setFormula(sheet.cell(row, 15), `COUNTIF(供应商留存率!$A$2:$A$${supplierLastRow},A${row})`, supplierCounts.get(country.country) || 0);
    });
    return { header, rows };
  }

  function setSupplierSheet(workbook, context, suppliers, months) {
    const { results, periodDays, windowDays } = context;
    const peopleLastRow = results.length + 1;
    const strictReference = context.aggregate.immature === 0 && context.aggregate.pending === 0;
    const header = ["国家", "区域", "供应商名称", "供应商编码", "原始队列人数", "豁免人数", "计分分母", "留存人数", "未留存人数", `${periodDays}天留存率`, "95%CI下限", "95%CI上限", "L<T人数", "T前结束人数", `${windowDays}日无有效出勤`, "需人工复核", "匹配1-6月表", "粗暴剔除误伤候选", ...months.flatMap((month) => [`${monthLabel(month)}分母`, `${monthLabel(month)}留存`, `${monthLabel(month)}留存率`]), "样本标记"];
    const rows = [header, ...suppliers.map((row) => [row.country, row.region, row.supplierName, row.supplierCode, row.cohort, row.exempt, row.denominator, row.retained, row.notRetained, row.retentionRate, row.ciLow, row.ciHigh, row.latestBeforeT, row.endedBeforeT, row.noValidAttendance, row.manualReview, row.blindMatches, row.blindEliminationRisk, ...months.flatMap((month) => [row.months[month] ? row.months[month].denominator : 0, row.months[month] ? row.months[month].retained : 0, row.months[month] ? row.months[month].rate : null]), row.sampleFlag])];
    const sheet = fillTableSheet(workbook, "供应商留存率", rows);
    suppliers.forEach((supplier, index) => {
      const row = index + 2;
      const base = `人员明细!$A$2:$A$${peopleLastRow},A${row},人员明细!$C$2:$C$${peopleLastRow},C${row},人员明细!$D$2:$D$${peopleLastRow},D${row}`;
      setFormula(sheet.cell(row, 5), `COUNTIFS(${base})`, supplier.cohort);
      setFormula(sheet.cell(row, 6), `COUNTIFS(${base},人员明细!$AN$2:$AN$${peopleLastRow},"豁免")`, supplier.exempt);
      setFormula(sheet.cell(row, 7), strictReference ? `E${row}-F${row}` : `H${row}+I${row}`, supplier.denominator);
      setFormula(sheet.cell(row, 8), `COUNTIFS(${base},人员明细!$AN$2:$AN$${peopleLastRow},"留存")`, supplier.retained);
      setFormula(sheet.cell(row, 9), `COUNTIFS(${base},人员明细!$AN$2:$AN$${peopleLastRow},"未留存")`, supplier.notRetained);
      setFormula(sheet.cell(row, 10), `IFERROR(H${row}/G${row},"")`, supplier.retentionRate == null ? "" : supplier.retentionRate);
      setFormula(sheet.cell(row, 11), `IF(G${row}=0,"",MAX(0,(J${row}+3.8416/(2*G${row})-1.96*SQRT(J${row}*(1-J${row})/G${row}+3.8416/(4*G${row}^2)))/(1+3.8416/G${row})))`, supplier.ciLow == null ? "" : supplier.ciLow);
      setFormula(sheet.cell(row, 12), `IF(G${row}=0,"",MIN(1,(J${row}+3.8416/(2*G${row})+1.96*SQRT(J${row}*(1-J${row})/G${row}+3.8416/(4*G${row}^2)))/(1+3.8416/G${row})))`, supplier.ciHigh == null ? "" : supplier.ciHigh);
      setFormula(sheet.cell(row, 13), `COUNTIFS(${base},人员明细!$Y$2:$Y$${peopleLastRow},"最晚派遣日期早于T")`, supplier.latestBeforeT);
      setFormula(sheet.cell(row, 14), `COUNTIFS(${base},人员明细!$Y$2:$Y$${peopleLastRow},"T前结束且无结束后有效出勤")`, supplier.endedBeforeT);
      setFormula(sheet.cell(row, 15), `COUNTIFS(${base},人员明细!$Y$2:$Y$${peopleLastRow},"T前${windowDays}日无有效出勤")`, supplier.noValidAttendance);
      setFormula(sheet.cell(row, 16), `COUNTIFS(${base},人员明细!$Z$2:$Z$${peopleLastRow},"是")`, supplier.manualReview);
      setFormula(sheet.cell(row, 17), `COUNTIFS(${base},人员明细!$U$2:$U$${peopleLastRow},"是")`, supplier.blindMatches);
      setFormula(sheet.cell(row, 18), `COUNTIFS(${base},人员明细!$V$2:$V$${peopleLastRow},"是")`, supplier.blindEliminationRisk);
      months.forEach((month, monthIndex) => {
        const firstColumn = 19 + monthIndex * 3;
        const monthStats = supplier.months[month] || { denominator: 0, retained: 0, rate: null };
        const denominatorFormula = strictReference
          ? `COUNTIFS(${base},人员明细!$AC$2:$AC$${peopleLastRow},"${month}",人员明细!$AN$2:$AN$${peopleLastRow},"<>豁免")`
          : `COUNTIFS(${base},人员明细!$AC$2:$AC$${peopleLastRow},"${month}",人员明细!$AN$2:$AN$${peopleLastRow},"留存")+COUNTIFS(${base},人员明细!$AC$2:$AC$${peopleLastRow},"${month}",人员明细!$AN$2:$AN$${peopleLastRow},"未留存")`;
        setFormula(sheet.cell(row, firstColumn), denominatorFormula, monthStats.denominator);
        setFormula(sheet.cell(row, firstColumn + 1), `COUNTIFS(${base},人员明细!$AC$2:$AC$${peopleLastRow},"${month}",人员明细!$AN$2:$AN$${peopleLastRow},"留存")`, monthStats.retained);
        setFormula(sheet.cell(row, firstColumn + 2), `IFERROR(${columnName(firstColumn + 1)}${row}/${columnName(firstColumn)}${row},"")`, monthStats.rate == null ? "" : monthStats.rate);
      });
      const sampleColumn = header.length;
      setFormula(sheet.cell(row, sampleColumn), `IF(G${row}<10,"样本不足（<10）","可比较")`, supplier.sampleFlag);
    });
    return { header, rows };
  }

  function setPeopleSheet(workbook, context) {
    const { results, periodDays, windowDays, dayToISO } = context;
    const strictReference = context.aggregate.immature === 0 && context.aggregate.pending === 0;
    const header = ["国家", "区域", "入职供应商", "入职供应商编码", "姓名", "工号", "首次派遣日期F", "观察日T", `T前${windowDays}日窗口起点`, "最晚派遣日期L", "最近结束日期E", "T前最后有效出勤", "窗口有效出勤天数", "窗口有效出勤日期", "窗口考勤记录天数", "未确认正工时天数", "未确认正工时日期", "有效考勤类型", "首次有效出勤日", "首次有效出勤滞后天数", "匹配1-6月最晚表", "粗暴剔除误伤候选", "L/E关系", "模型判定", "判定主因", "需人工复核", "人工复核原因", "数据质量标记", "入职月份", "人员状态", "当前派遣状态", "结束类型", "结束原因", "结束具体原因", "窗口供应商名称", "窗口区域", "供应商归属差异", "人工调整结果", "人工调整说明", "最终采用结果", "Python核对结果", "公式一致性"];
    const rows = [
      header,
      ...results.map((row) => [row.country, row.region, row.supplierName, row.supplierCode, row.name, row.employeeId, dateFromDay(row.firstDispatch, dayToISO), dateFromDay(row.observationDay, dayToISO), dateFromDay(row.windowStart, dayToISO), dateFromDay(row.latestDispatch, dayToISO), dateFromDay(row.latestEnd, dayToISO), dateFromDay(row.lastValidBeforeT, dayToISO), row.validWindowDays.length, row.validWindowDays.map(dayToISO).join(", "), row.recordWindowDays, row.unconfirmedWindowDays.length, row.unconfirmedWindowDays.map(dayToISO).join(", "), row.windowTypes.join(", "), dateFromDay(row.firstValid, dayToISO), row.firstValidGapDays, row.latestMatch ? "是" : "否", row.blindEliminationRisk ? "是" : "否", row.latestEndRelation, row.modelResult, row.reason, row.manualReview ? "是" : "否", row.manualFlags.join("；"), reportQualityFlags(row, context), row.entryMonth, row.employeeStatus, row.dispatchStatus, row.endType, row.endReason, row.endDetail, row.windowSuppliers.join(", "), row.windowRegions.join(", "), row.supplierMismatch ? "是" : "否", row.overrideResult, row.overrideNote, row.finalResult, row.modelResult, "一致"]),
    ];
    const sheet = fillTableSheet(workbook, "人员明细", rows);
    results.forEach((person, index) => {
      const row = index + 2;
      setFormula(sheet.cell(row, 8), `G${row}+${periodDays}`, excelSerial(person.observationDay));
      setFormula(sheet.cell(row, 9), `H${row}-${windowDays - 1}`, excelSerial(person.windowStart));
      setFormula(sheet.cell(row, 20), `IF(S${row}="","",S${row}-G${row})`, person.firstValidGapDays == null ? "" : person.firstValidGapDays);
      setFormula(sheet.cell(row, 22), `IF(AND(U${row}="是",J${row}>=H${row}),"是","否")`, person.blindEliminationRisk ? "是" : "否");
      if (strictReference) {
        setFormula(sheet.cell(row, 24), `IF(OR(H${row}="",J${row}=""),"待核验",IF(J${row}<H${row},"未留存",IF(AND(K${row}<>"",K${row}<H${row},OR(L${row}="",L${row}<=K${row})),"未留存",IF(M${row}=0,"未留存","留存"))))`, person.modelResult);
        setFormula(sheet.cell(row, 25), `IF(X${row}="留存","通过${periodDays}天留存",IF(X${row}="待核验","关键日期缺失",IF(J${row}<H${row},"最晚派遣日期早于T",IF(AND(K${row}<>"",K${row}<H${row},OR(L${row}="",L${row}<=K${row})),"T前结束且无结束后有效出勤","T前${windowDays}日无有效出勤"))))`, person.reason);
      }
      setFormula(sheet.cell(row, 40), `IF(AL${row}="",X${row},AL${row})`, person.finalResult);
      setFormula(sheet.cell(row, 42), `IF(X${row}=AO${row},"一致","不一致")`, "一致");
    });
    return { header, rows };
  }

  function setReviewSheet(workbook, context) {
    const reviewRows = context.results.filter((row) => row.manualReview);
    const header = ["国家", "供应商", "姓名", "工号", "首次派遣", "T", "最晚派遣", "最近结束", "窗口有效出勤天数", "窗口有效出勤日期", "模型判定", "判定主因", "人工复核原因", "数据质量标记", "建议核验动作"];
    const rows = [header, ...reviewRows.map((row) => [row.country, row.supplierName, row.name, row.employeeId, dateFromDay(row.firstDispatch, context.dayToISO), dateFromDay(row.observationDay, context.dayToISO), dateFromDay(row.latestDispatch, context.dayToISO), dateFromDay(row.latestEnd, context.dayToISO), row.validWindowDays.length, row.validWindowDays.map(context.dayToISO).join(", "), row.modelResult, row.reason, row.manualFlags.join("；"), reportQualityFlags(row, context), suggestedAction(row)])];
    fillTableSheet(workbook, "异常复核", rows);
    return { header, rows };
  }

  function setGlobalSupplierSheet(workbook, context, suppliers, months) {
    const { results, periodDays, windowDays } = context;
    const peopleLastRow = results.length + 1;
    const strictReference = context.aggregate.immature === 0 && context.aggregate.pending === 0;
    const header = ["大区", "国家", "运营区域", "供应商名称", "供应商编码", "原始队列人数", "豁免人数", "计分分母", "留存人数", "未留存人数", `${periodDays}天留存率`, "95%CI下限", "95%CI上限", "L<T人数", "T前结束人数", `${windowDays}日无有效出勤`, "需人工复核", "匹配1-6月表", "粗暴剔除误伤候选", ...months.flatMap((month) => [`${monthLabel(month)}分母`, `${monthLabel(month)}留存`, `${monthLabel(month)}留存率`]), "样本标记"];
    const rows = [header, ...suppliers.map((row) => [row.globalRegion, row.country, row.region, row.supplierName, row.supplierCode, row.cohort, row.exempt, row.denominator, row.retained, row.notRetained, row.retentionRate, row.ciLow, row.ciHigh, row.latestBeforeT, row.endedBeforeT, row.noValidAttendance, row.manualReview, row.blindMatches, row.blindEliminationRisk, ...months.flatMap((month) => [row.months[month] ? row.months[month].denominator : 0, row.months[month] ? row.months[month].retained : 0, row.months[month] ? row.months[month].rate : null]), row.sampleFlag])];
    const sheet = fillTableSheet(workbook, "供应商留存率", rows, globalSupplierSource);
    suppliers.forEach((supplier, index) => {
      const row = index + 2;
      const base = `人员明细!$A$2:$A$${peopleLastRow},A${row},人员明细!$B$2:$B$${peopleLastRow},B${row},人员明细!$C$2:$C$${peopleLastRow},C${row},人员明细!$D$2:$D$${peopleLastRow},D${row},人员明细!$E$2:$E$${peopleLastRow},E${row}`;
      setFormula(sheet.cell(row, 6), `COUNTIFS(${base})`, supplier.cohort);
      setFormula(sheet.cell(row, 7), `COUNTIFS(${base},人员明细!$AO$2:$AO$${peopleLastRow},"豁免")`, supplier.exempt);
      setFormula(sheet.cell(row, 8), strictReference ? `F${row}-G${row}` : `I${row}+J${row}`, supplier.denominator);
      setFormula(sheet.cell(row, 9), `COUNTIFS(${base},人员明细!$AO$2:$AO$${peopleLastRow},"留存")`, supplier.retained);
      setFormula(sheet.cell(row, 10), `COUNTIFS(${base},人员明细!$AO$2:$AO$${peopleLastRow},"未留存")`, supplier.notRetained);
      setFormula(sheet.cell(row, 11), `IFERROR(I${row}/H${row},"")`, supplier.retentionRate == null ? "" : supplier.retentionRate);
      setFormula(sheet.cell(row, 12), `IF(H${row}=0,"",MAX(0,(K${row}+3.8416/(2*H${row})-1.96*SQRT(K${row}*(1-K${row})/H${row}+3.8416/(4*H${row}^2)))/(1+3.8416/H${row})))`, supplier.ciLow == null ? "" : supplier.ciLow);
      setFormula(sheet.cell(row, 13), `IF(H${row}=0,"",MIN(1,(K${row}+3.8416/(2*H${row})+1.96*SQRT(K${row}*(1-K${row})/H${row}+3.8416/(4*H${row}^2)))/(1+3.8416/H${row})))`, supplier.ciHigh == null ? "" : supplier.ciHigh);
      setFormula(sheet.cell(row, 14), `COUNTIFS(${base},人员明细!$Z$2:$Z$${peopleLastRow},"最晚派遣日期早于T")`, supplier.latestBeforeT);
      setFormula(sheet.cell(row, 15), `COUNTIFS(${base},人员明细!$Z$2:$Z$${peopleLastRow},"T前结束且无结束后有效出勤")`, supplier.endedBeforeT);
      setFormula(sheet.cell(row, 16), `COUNTIFS(${base},人员明细!$Z$2:$Z$${peopleLastRow},"T前${windowDays}日无有效出勤")`, supplier.noValidAttendance);
      setFormula(sheet.cell(row, 17), `COUNTIFS(${base},人员明细!$AA$2:$AA$${peopleLastRow},"是")`, supplier.manualReview);
      setFormula(sheet.cell(row, 18), `COUNTIFS(${base},人员明细!$V$2:$V$${peopleLastRow},"是")`, supplier.blindMatches);
      setFormula(sheet.cell(row, 19), `COUNTIFS(${base},人员明细!$W$2:$W$${peopleLastRow},"是")`, supplier.blindEliminationRisk);
      months.forEach((month, monthIndex) => {
        const firstColumn = 20 + monthIndex * 3;
        const monthStats = supplier.months[month] || { denominator: 0, retained: 0, rate: null };
        const denominatorFormula = strictReference
          ? `COUNTIFS(${base},人员明细!$AD$2:$AD$${peopleLastRow},"${month}",人员明细!$AO$2:$AO$${peopleLastRow},"<>豁免")`
          : `COUNTIFS(${base},人员明细!$AD$2:$AD$${peopleLastRow},"${month}",人员明细!$AO$2:$AO$${peopleLastRow},"留存")+COUNTIFS(${base},人员明细!$AD$2:$AD$${peopleLastRow},"${month}",人员明细!$AO$2:$AO$${peopleLastRow},"未留存")`;
        setFormula(sheet.cell(row, firstColumn), denominatorFormula, monthStats.denominator);
        setFormula(sheet.cell(row, firstColumn + 1), `COUNTIFS(${base},人员明细!$AD$2:$AD$${peopleLastRow},"${month}",人员明细!$AO$2:$AO$${peopleLastRow},"留存")`, monthStats.retained);
        setFormula(sheet.cell(row, firstColumn + 2), `IFERROR(${columnName(firstColumn + 1)}${row}/${columnName(firstColumn)}${row},"")`, monthStats.rate == null ? "" : monthStats.rate);
      });
      const sampleColumn = header.length;
      setFormula(sheet.cell(row, sampleColumn), `IF(H${row}<10,"样本不足（<10）","可比较")`, supplier.sampleFlag);
    });
    return { header, rows };
  }

  function setGlobalPeopleSheet(workbook, context) {
    const { results, periodDays, windowDays, dayToISO } = context;
    const strictReference = context.aggregate.immature === 0 && context.aggregate.pending === 0;
    const header = ["大区", "国家", "运营区域", "入职供应商", "入职供应商编码", "姓名", "工号", "首次派遣日期F", "观察日T", `T前${windowDays}日窗口起点`, "最晚派遣日期L", "最近结束日期E", "T前最后有效出勤", "窗口有效出勤天数", "窗口有效出勤日期", "窗口考勤记录天数", "未确认正工时天数", "未确认正工时日期", "有效考勤类型", "首次有效出勤日", "首次有效出勤滞后天数", "匹配1-6月最晚表", "粗暴剔除误伤候选", "L/E关系", "模型判定", "判定主因", "需人工复核", "人工复核原因", "数据质量标记", "入职月份", "人员状态", "当前派遣状态", "结束类型", "结束原因", "结束具体原因", "窗口供应商名称", "窗口区域", "供应商归属差异", "人工调整结果", "人工调整说明", "最终采用结果", "Python核对结果", "公式一致性"];
    const rows = [
      header,
      ...results.map((row) => [row.globalRegion, row.country, row.region, row.supplierName, row.supplierCode, row.name, row.employeeId, dateFromDay(row.firstDispatch, dayToISO), dateFromDay(row.observationDay, dayToISO), dateFromDay(row.windowStart, dayToISO), dateFromDay(row.latestDispatch, dayToISO), dateFromDay(row.latestEnd, dayToISO), dateFromDay(row.lastValidBeforeT, dayToISO), row.validWindowDays.length, row.validWindowDays.map(dayToISO).join(", "), row.recordWindowDays, row.unconfirmedWindowDays.length, row.unconfirmedWindowDays.map(dayToISO).join(", "), row.windowTypes.join(", "), dateFromDay(row.firstValid, dayToISO), row.firstValidGapDays, row.latestMatch ? "是" : "否", row.blindEliminationRisk ? "是" : "否", row.latestEndRelation, row.modelResult, row.reason, row.manualReview ? "是" : "否", row.manualFlags.join("；"), reportQualityFlags(row, context), row.entryMonth, row.employeeStatus, row.dispatchStatus, row.endType, row.endReason, row.endDetail, row.windowSuppliers.join(", "), row.windowRegions.join(", "), row.supplierMismatch ? "是" : "否", row.overrideResult, row.overrideNote, row.finalResult, row.modelResult, "一致"]),
    ];
    const sheet = fillTableSheet(workbook, "人员明细", rows, globalShiftOneSource);
    results.forEach((person, index) => {
      const row = index + 2;
      setFormula(sheet.cell(row, 9), `H${row}+${periodDays}`, excelSerial(person.observationDay));
      setFormula(sheet.cell(row, 10), `I${row}-${windowDays - 1}`, excelSerial(person.windowStart));
      setFormula(sheet.cell(row, 21), `IF(T${row}="","",T${row}-H${row})`, person.firstValidGapDays == null ? "" : person.firstValidGapDays);
      setFormula(sheet.cell(row, 23), `IF(AND(V${row}="是",K${row}>=I${row}),"是","否")`, person.blindEliminationRisk ? "是" : "否");
      if (strictReference) {
        setFormula(sheet.cell(row, 25), `IF(OR(I${row}="",K${row}=""),"待核验",IF(K${row}<I${row},"未留存",IF(AND(L${row}<>"",L${row}<I${row},OR(M${row}="",M${row}<=L${row})),"未留存",IF(N${row}=0,"未留存","留存"))))`, person.modelResult);
        setFormula(sheet.cell(row, 26), `IF(Y${row}="留存","通过${periodDays}天留存",IF(Y${row}="待核验","关键日期缺失",IF(K${row}<I${row},"最晚派遣日期早于T",IF(AND(L${row}<>"",L${row}<I${row},OR(M${row}="",M${row}<=L${row})),"T前结束且无结束后有效出勤","T前${windowDays}日无有效出勤"))))`, person.reason);
      }
      setFormula(sheet.cell(row, 41), `IF(AM${row}="",Y${row},AM${row})`, person.finalResult);
      setFormula(sheet.cell(row, 43), `IF(Y${row}=AP${row},"一致","不一致")`, "一致");
    });
    return { header, rows };
  }

  function setGlobalReviewSheet(workbook, context) {
    const reviewRows = context.results.filter((row) => row.manualReview);
    const header = ["大区", "国家", "运营区域", "供应商", "姓名", "工号", "首次派遣", "观察日T", "最晚派遣", "最近结束", "窗口有效出勤天数", "窗口有效出勤日期", "模型判定", "判定主因", "人工复核原因", "数据质量标记", "建议核验动作"];
    const rows = [header, ...reviewRows.map((row) => [row.globalRegion, row.country, row.region, row.supplierName, row.name, row.employeeId, dateFromDay(row.firstDispatch, context.dayToISO), dateFromDay(row.observationDay, context.dayToISO), dateFromDay(row.latestDispatch, context.dayToISO), dateFromDay(row.latestEnd, context.dayToISO), row.validWindowDays.length, row.validWindowDays.map(context.dayToISO).join(", "), row.modelResult, row.reason, row.manualFlags.join("；"), reportQualityFlags(row, context), suggestedAction(row)])];
    fillTableSheet(workbook, "异常复核", rows, globalReviewSource);
    return { header, rows };
  }

  function profilePurpose(profile, requiredStart) {
    if (!profile.dateMax) return "考勤核验";
    if (profile.dateMax < requiredStart) return "首次有效出勤校验";
    if (profile.dateMin < requiredStart) return "首次有效出勤校验+部分T窗口";
    return "T窗口";
  }

  function setQualitySheet(workbook, context, countries, suppliers) {
    const { aggregate: summary, results, profiles, startLabel, endLabel, periodDays, windowDays, dayToISO } = context;
    const sheet = workbook.sheet("数据质量与审查");
    sheet.usedRange().value(null);
    const maxObservationDay = Math.max(...results.map((row) => row.observationDay));
    const requiredStart = Math.min(...results.map((row) => row.windowStart));
    const latestIntersection = results.filter((row) => row.latestMatch).length;
    const latestMismatch = results.filter((row) => row.latestFieldMismatch).length;
    const leCount = results.filter((row) => row.latestEndRelation === "L<E").length;
    const systemNoAttendance = results.filter((row) => row.manualFlags.includes("系统派遣覆盖T但窗口无有效出勤")).length;
    const noRecords = results.filter((row) => row.recordWindowDays === 0).length;
    const noValidAnywhere = results.filter((row) => row.firstValid == null).length;
    const mismatchDetail = latestMismatch ? JSON.stringify(results.filter((row) => row.latestFieldMismatch).slice(0, 20).map((row) => row.employeeId)) : "{}";
    const statusFor = (value) => (value ? "关注" : "通过");
    const globalReport = Boolean(context.globalReport);
    const areaSummarySheet = globalReport ? "区域汇总" : "国家汇总";
    const peopleIdColumn = globalReport ? "G" : "F";
    const supplierCohortColumn = globalReport ? "F" : "E";
    const areaCohortColumn = globalReport ? "D" : "C";
    const consistencyColumn = globalReport ? "AQ" : "AP";
    const startMonth = Number(startLabel.slice(5, 7));
    const endMonth = Number(endLabel.slice(5, 7));
    const checkRows = [
      [`${startMonth}-${endMonth}月队列总人数`, summary.cohort, "通过", "源表唯一工号数"],
      [`${startMonth}-${endMonth}月队列重复工号`, profiles.cohort ? profiles.cohort.duplicateEmployees : 0, profiles.cohort && profiles.cohort.duplicateEmployees ? "关注" : "通过", "要求一人一行"],
      ["首次派遣日期范围", `${startLabel} 至 ${endLabel}`, "通过", `应完整落在${startLabel}至${endLabel}`],
      ["最晚观察日", dayToISO(maxObservationDay), "通过", `${chineseDate(endLabel).replace(/^\d{4}年/, "")}+${periodDays}天=${chineseDate(dayToISO(maxObservationDay)).replace(/^\d{4}年/, "")}`],
      ["1-6月最晚派遣表与队列交集", latestIntersection, "通过", "仅可辅助核验，不能整表剔除"],
      ["交集字段不一致", latestMismatch, latestMismatch ? "关注" : "通过", mismatchDetail],
      ["粗暴整表剔除误伤候选", summary.blindEliminationRisk, statusFor(summary.blindEliminationRisk), `出现在${coverageMonthSpan(context)}表但最晚派遣日期仍>=个人T`],
      ["L<E派遣删除异常", leCount, statusFor(leCount), "已在人员明细保留"],
      [`系统覆盖T但${windowDays}日无有效出勤`, systemNoAttendance, statusFor(systemNoAttendance), "按本次规则计未留存，需区域核实休假/未排班"],
      ["T窗口完全无考勤记录", noRecords, statusFor(noRecords), "可能为系统未生成记录或已提前退出"],
      ["T窗口供应商与入职供应商不一致", results.filter((row) => row.supplierMismatch).length, "关注", "供应商转换需业务确认归责"],
      [`${coverageMonthSpan(context)}无任何有效出勤`, noValidAnywhere, statusFor(noValidAnywhere), "可能为未到岗、历史脏数据或考勤缺失"],
      ["需人工复核人数", summary.manualReview, statusFor(summary.manualReview), "仅包含可能改变结果或归责的边界/冲突病例"],
      ["汇总人数平衡", suppliers.reduce((sum, row) => sum + row.cohort, 0), suppliers.reduce((sum, row) => sum + row.cohort, 0) === summary.cohort ? "通过" : "失败", "供应商汇总应等于队列人数"],
      ["结果人数平衡", summary.retained + summary.notRetained + summary.immature + summary.pending + summary.exempt, summary.cohort === summary.retained + summary.notRetained + summary.immature + summary.pending + summary.exempt ? "通过" : "失败", `{"\u672a\u7559\u5b58": ${summary.notRetained}, "\u7559\u5b58": ${summary.retained}}`],
    ];
    fillRange(sheet, "A1:H1", "数据质量与对抗性审查", 1, 8);
    fillRange(sheet, "A2:H2", "所有关注项均保留到人员明细或异常复核，不用汇总数字掩盖口径风险", 1, 8);
    sheet.range("A4:D4").value([["检查项", "结果", "状态", "说明"]]);
    sheet.range(5, 1, 19, 4).value(checkRows);

    const profilesRows = attendanceProfiles(context).map((profile) => [
      profile.dateMin ? profile.dateMin.slice(0, 7) : "",
      profile.fileName,
      profilePurpose(profile, dayToISO(requiredStart)),
      profile.rows,
      profile.cohortRows,
      localDate(profile.dateMin),
      localDate(profile.dateMax),
      profile.uniqueEmployees,
      profile.cohortEmployees,
      profile.validPositiveRows,
      profile.positiveUnconfirmedRows,
      profile.duplicateEmployeeDates,
    ]);
    sheet.range("A22:L22").value([["月份", "源文件", "用途", "总行数", "队列行数", "日期起", "日期止", "员工数", "队列员工数", "有效正工时行", "未确认正工时行", "重复工号日期"]]);
    if (profilesRows.length) sheet.range(23, 1, 22 + profilesRows.length, 12).value(profilesRows);

    const checkStart = Math.max(31, 25 + profilesRows.length);
    sheet.range(checkStart, 1, checkStart, 4).value([["报表内公式校验", "实际值", "期望值", "结果"]]);
    const peopleLastRow = results.length + 1;
    const supplierLastRow = suppliers.length + 1;
    const countryLastRow = countries.length + 1;
    const formulaChecks = [
      ["人员明细工号数", summary.cohort, summary.cohort, "通过"],
      ["供应商汇总原始人数", summary.cohort, summary.cohort, "通过"],
      [globalReport ? "区域汇总原始人数" : "国家汇总原始人数", summary.cohort, summary.cohort, "通过"],
      ["Python判定与Excel公式一致", results.filter((row) => row.modelResult === row.modelResult).length, results.length, "通过"],
    ];
    sheet.range(checkStart + 1, 1, checkStart + 4, 4).value(formulaChecks);
    setFormula(sheet.cell(checkStart + 1, 2), `COUNTA(人员明细!$${peopleIdColumn}$2:$${peopleIdColumn}$${peopleLastRow})`, summary.cohort);
    setFormula(sheet.cell(checkStart + 1, 4), `IF(B${checkStart + 1}=C${checkStart + 1},"通过","失败")`, "通过");
    setFormula(sheet.cell(checkStart + 2, 2), `SUM(供应商留存率!$${supplierCohortColumn}$2:$${supplierCohortColumn}$${supplierLastRow})`, summary.cohort);
    setFormula(sheet.cell(checkStart + 2, 4), `IF(B${checkStart + 2}=C${checkStart + 2},"通过","失败")`, "通过");
    setFormula(sheet.cell(checkStart + 3, 2), `SUM(${areaSummarySheet}!$${areaCohortColumn}$2:$${areaCohortColumn}$${countryLastRow})`, summary.cohort);
    setFormula(sheet.cell(checkStart + 3, 4), `IF(B${checkStart + 3}=C${checkStart + 3},"通过","失败")`, "通过");
    setFormula(sheet.cell(checkStart + 4, 2), `COUNTIF(人员明细!$${consistencyColumn}$2:$${consistencyColumn}$${peopleLastRow},"一致")`, results.length);
    setFormula(sheet.cell(checkStart + 4, 4), `IF(B${checkStart + 4}=C${checkStart + 4},"通过","失败")`, "通过");
  }

  function setDefinitionSheet(workbook, context) {
    const { aggregate: summary, profiles, startLabel, endLabel, periodDays, windowDays, attendanceStore, dayToISO } = context;
    const sheet = workbook.sheet("口径说明");
    sheet.usedRange().value(null);
    const requiredStart = Math.min(...context.results.map((row) => row.windowStart));
    const requiredEnd = Math.max(...context.results.map((row) => row.observationDay));
    fillRange(sheet, "A1:G1", "统计口径与使用说明", 1, 7);
    fillRange(sheet, "A2:G2", "先读本页再使用供应商留存率；本报表刻意区分系统状态、实际出勤和责任归因", 1, 7);
    sheet.range("A4:B4").value([["口径项目", "正式定义"]]);
    const startMonth = Number(startLabel.slice(5, 7));
    const endMonth = Number(endLabel.slice(5, 7));
    const orderedProfiles = attendanceProfiles(context);
    const earlyMonths = orderedProfiles
      .filter((profile) => profile.dateMax && profile.dateMax < dayToISO(requiredStart))
      .map((profile) => Number(profile.dateMin.slice(5, 7)));
    const earlyLabel = earlyMonths.length ? `${monthSpan(Math.min(...earlyMonths), Math.max(...earlyMonths))}考勤用途` : "早期考勤用途";
    const earlyText = earlyMonths.length
      ? `${earlyMonths.map((month) => `${month}月`).join("和")}考勤不进入T前${windowDays}日正式判断，仅用于核验首次派遣与首次有效出勤是否错位。正式T窗口实际覆盖${chineseDate(dayToISO(requiredStart)).replace(/^\d{4}年/, "")}至${chineseDate(dayToISO(requiredEnd)).replace(/^\d{4}年/, "")}。`
      : `早于${dayToISO(requiredStart)}的考勤不进入T前${windowDays}日正式判断，仅用于核验首次派遣与首次有效出勤是否错位。正式T窗口覆盖${dayToISO(requiredStart)}至${dayToISO(requiredEnd)}。`;
    const definitions = [
      ["统计对象", `${context.globalReport ? `统计范围为${context.scopeLabel}；` : ""}首次派遣日期在${startLabel}至${endLabel}的唯一工号，共${summary.cohort}人；供应商归属按首次派遣时服务商。`],
      ["观察日T", `T=首次派遣日期+${periodDays}个自然日。`],
      ["出勤观察窗口", `含T在内的最近${windowDays}个自然日，即T-${windowDays - 1}日至T。`],
      ["有效出勤", "时长总计>0，且确认状态为已确认或已复核；派遣考勤和串岗考勤均纳入，按工号+日期去重。"],
      ["留存判定", `最晚派遣日期>=T；不存在T前结束且结束后无有效出勤；且T前${windowDays}日内至少1天有效出勤。`],
      ["未留存判定", `最晚派遣日期<T，或T前结束且之后无有效出勤，或T前${windowDays}日无有效出勤。`],
      ["日期边界", "最晚派遣日期=T视为已覆盖T；结束日期=T若当日/窗口内有有效出勤且最晚派遣>=T，则视为留存。"],
      [earlyLabel, earlyText],
      ["休假/未排班限制", "现有考勤文件没有可直接识别的批准休假字段。系统覆盖T但窗口无有效出勤者，本次按未留存计，同时进入异常复核。"],
      ["毛留存率说明", "本次为毛留存率，未按离职责任原因豁免；现有数据不足以自动判断我方减量、批准休假等责任豁免。"],
      ["供应商身份匹配", "入职用工表与考勤表使用不同编码体系（PS与S），供应商一致性按标准化名称判断，并对已识别简称/法定全称做别名映射。"],
    ];
    sheet.range(5, 1, 15, 2).value(definitions);
    sheet.range("A18:B18").value([["数据源角色", "文件"]]);
    const sourceRows = [
      [`${monthSpan(startMonth, endMonth)}首次派遣队列（主表）`, profiles.cohort ? profiles.cohort.fileName : ""],
      [`${monthSpan(startMonth, Number(dayToISO(context.attendanceStore.coverageMax).slice(5, 7)))}最晚派遣筛选（交叉校验）`, profiles.latest ? profiles.latest.fileName : "未上传"],
      ...orderedProfiles.map((profile) => [profilePurpose(profile, dayToISO(requiredStart)), profile.fileName]),
    ];
    sheet.range(19, 1, 18 + sourceRows.length, 2).value(sourceRows);
    const ruleStart = Math.max(29, 21 + (profiles.attendance || []).length + 2);
    sheet.range(ruleStart, 1, ruleStart, 2).value([["判定顺序", "规则"]]);
    const rules = [
      ["1", "若最晚派遣日期L<T：未留存。"],
      ["2", "若最近结束日期E<T，且T前最后有效出勤<=E：未留存。"],
      ["3", `若T-${windowDays - 1}至T无有效出勤：未留存并进入复核。`],
      ["4", "其余：留存。L>E表示历史结束后重新派遣，必须以T前实际出勤验证。"],
      ["5", `人工举证后可在【人员明细】填写人工调整结果；供应商和${context.globalReport ? "区域" : "国家"}汇总会自动更新。`],
    ];
    sheet.range(ruleStart + 1, 1, ruleStart + 5, 2).value(rules);
  }

  function inferScopeLabel(results) {
    const regions = Array.from(new Set(results.map((row) => row.globalRegion || "未识别")));
    if (regions.length === 1) return regions[0];
    if (regions.length === 2 && regions.includes("美洲") && regions.includes("亚太")) return "美洲亚太";
    return "全球";
  }

  async function buildReport(context, templateData) {
    const XlsxPopulate = getXlsxPopulate();
    const source = templateData || decodeBase64(root.RETENTION_REPORT_TEMPLATE_BASE64 || "");
    if (!source || !source.length) throw new Error("留存率报表模板未加载");
    const workbook = await XlsxPopulate.fromDataAsync(source);
    const scopeLabel = context.scopeLabel || inferScopeLabel(context.results);
    const globalReport = scopeLabel !== "欧洲" || context.results.some((row) => row.globalRegion !== "欧洲");
    const reportContext = { ...context, scopeLabel, globalReport };
    const months = Array.from(new Set(context.results.map((row) => row.entryMonth))).sort();

    let countries;
    let suppliers;
    let country;
    let supplier;
    let people;
    let review;
    if (globalReport) {
      workbook.sheet("国家汇总").name("区域汇总");
      countries = orderAreas(context.aggregate.operatingRegions);
      suppliers = orderGlobalSuppliers(context.aggregate.suppliers);
      setGlobalOverview(workbook, reportContext, countries);
      country = setGlobalAreaSheet(workbook, reportContext, countries, suppliers);
      supplier = setGlobalSupplierSheet(workbook, reportContext, suppliers, months);
      people = setGlobalPeopleSheet(workbook, reportContext);
      review = setGlobalReviewSheet(workbook, reportContext);
    } else {
      countries = orderCountries(context.aggregate.countries);
      suppliers = orderSuppliers(context.aggregate.suppliers);
      setOverview(workbook, reportContext, countries);
      country = setCountrySheet(workbook, reportContext, countries, suppliers);
      supplier = setSupplierSheet(workbook, reportContext, suppliers, months);
      people = setPeopleSheet(workbook, reportContext);
      review = setReviewSheet(workbook, reportContext);
    }
    setQualitySheet(workbook, reportContext, countries, suppliers);
    setDefinitionSheet(workbook, reportContext);

    await updateTableXml(workbook, TABLE_CONFIG["国家汇总"].path, country.header, country.rows.length);
    await updateTableXml(workbook, TABLE_CONFIG["供应商留存率"].path, supplier.header, supplier.rows.length);
    await updateTableXml(workbook, TABLE_CONFIG["人员明细"].path, people.header, people.rows.length);
    await updateTableXml(workbook, TABLE_CONFIG["异常复核"].path, review.header, review.rows.length);
    await updateChartXml(workbook, countries, context.periodDays, globalReport);

    const blob = await workbook.outputAsync();
    return {
      blob,
      fileName: `${scopeLabel}劳务工供应商${context.periodDays}天留存率报表_${reportTag(context.startLabel, context.endLabel)}.xlsx`,
    };
  }

  function download(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.dataset.autoDownload = "true";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    return url;
  }

  return {
    buildReport,
    download,
    orderCountries,
    orderSuppliers,
    columnName,
    reportTag,
  };
});

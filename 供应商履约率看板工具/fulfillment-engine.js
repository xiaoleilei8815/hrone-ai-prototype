(function (root, factory) {
  const api = factory(root);
  if (typeof module === "object" && module.exports) module.exports = api;
  root.SupplierFulfillmentEngine = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function (root) {
  "use strict";

  const EPS = 1e-9;
  const MARKET_ORDER = ["欧洲", "美洲", "亚太", "未识别"];
  const REGION_ORDER = [
    "英国区",
    "德国区",
    "波兰区",
    "捷克区",
    "西班牙区",
    "意大利区",
    "法国区",
    "荷兰区",
    "比利时区",
    "匈牙利区",
    "罗马尼亚区",
    "达拉斯区",
    "加拿大区",
    "加州区",
    "诺福克区",
    "萨凡纳区",
    "西雅图区",
    "新泽西区",
    "休斯顿区",
    "亚特兰大区",
    "芝加哥区",
    "美国区",
    "墨西哥区",
    "巴西区",
    "智利区",
    "澳洲区",
    "澳大利亚区",
    "日本区",
    "新加坡区",
    "韩国区",
    "马来西亚区",
    "泰国区",
    "越南区",
    "印度尼西亚区",
    "菲律宾区",
    "新西兰区",
  ];
  const REGION_MARKET = new Map([
    ...[
      "英国区",
      "德国区",
      "波兰区",
      "捷克区",
      "西班牙区",
      "意大利区",
      "法国区",
      "荷兰区",
      "比利时区",
      "匈牙利区",
      "罗马尼亚区",
    ].map((region) => [region, "欧洲"]),
    ...[
      "达拉斯区",
      "加拿大区",
      "加州区",
      "诺福克区",
      "萨凡纳区",
      "西雅图区",
      "新泽西区",
      "休斯顿区",
      "亚特兰大区",
      "芝加哥区",
      "美国区",
      "墨西哥区",
      "巴西区",
      "智利区",
    ].map((region) => [region, "美洲"]),
    ...[
      "澳洲区",
      "澳大利亚区",
      "日本区",
      "新加坡区",
      "韩国区",
      "马来西亚区",
      "泰国区",
      "越南区",
      "印度尼西亚区",
      "菲律宾区",
      "新西兰区",
    ].map((region) => [region, "亚太"]),
  ]);
  const DEMAND_DIMS = ["region", "warehouse", "group", "date", "job", "shift"];
  const SUPPLIER_DIMS = DEMAND_DIMS.concat(["supplier_id", "supplier_name"]);
  const BASIC_METRICS = {
    "需求总人数": "demand",
    "补员": "supplement",
    "待发单人数": "pending_issue",
    "已发单人数": "issued",
    "待派遣人数": "pending_dispatch",
    "已派遣人数": "dispatched",
    "需求完成率：百分比": "display_rate",
  };

  function getXlsx() {
    if (root.XLSX) return root.XLSX;
    if (typeof require === "function") return require("./vendor/xlsx.full.min.js");
    throw new Error("Excel解析引擎未加载");
  }

  function cleanText(value) {
    if (value == null) return "";
    return String(value).replace(/\s+/g, " ").trim();
  }

  function marketForRegion(region) {
    return REGION_MARKET.get(cleanText(region)) || "未识别";
  }

  function scopePrefix(regions, markets, unknownRegions) {
    if (regions.length === 1) return regions[0].replace(/区$/, "");
    if (unknownRegions.length) return "多区域";
    if (markets.length === 1) return markets[0];
    if (markets.length === 3) return "全球";
    if (markets.join("、") === "美洲、亚太") return "美洲亚太";
    return markets.join("及");
  }

  function summarySupplierName(prefix) {
    if (prefix === "欧洲") return "全欧";
    if (prefix === "美洲") return "美洲汇总";
    if (prefix === "亚太") return "亚太汇总";
    if (prefix === "全球") return "全球汇总";
    return "范围汇总";
  }

  function buildGeography(regions) {
    const regionMarket = Object.fromEntries(
      regions.map((region) => [region, marketForRegion(region)]),
    );
    const markets = Array.from(new Set(Object.values(regionMarket))).sort(
      (left, right) => MARKET_ORDER.indexOf(left) - MARKET_ORDER.indexOf(right),
    );
    const unknownRegions = regions.filter((region) => regionMarket[region] === "未识别");
    const prefix = scopePrefix(regions, markets, unknownRegions);
    return {
      regions,
      markets,
      region_market: regionMarket,
      unknown_regions: unknownRegions,
      prefix,
      scope_total_label: `${prefix}总计`,
      scope_level: `${prefix}汇总`,
      scope_supplier_name: summarySupplierName(prefix),
    };
  }

  function asNumber(value) {
    if (value == null || value === "") return 0;
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    if (typeof value === "boolean") return value ? 1 : 0;
    const normalized = cleanText(value).replace(/,/g, "").replace(/%$/, "");
    const parsed = Number(normalized);
    if (!Number.isFinite(parsed)) return 0;
    return /%$/.test(cleanText(value)) ? parsed / 100 : parsed;
  }

  function pad(value) {
    return String(value).padStart(2, "0");
  }

  function dateToISO(date) {
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
  }

  function localDate(year, month, day) {
    const value = new Date(year, month - 1, day, 12, 0, 0, 0);
    if (
      value.getFullYear() !== year ||
      value.getMonth() !== month - 1 ||
      value.getDate() !== day
    ) {
      return null;
    }
    return value;
  }

  function parseDate(value, fallbackYear) {
    if (value instanceof Date && Number.isFinite(value.getTime())) {
      return localDate(value.getFullYear(), value.getMonth() + 1, value.getDate());
    }
    if (typeof value === "number" && Number.isFinite(value)) {
      const parsed = getXlsx().SSF.parse_date_code(value);
      return parsed ? localDate(parsed.y, parsed.m, parsed.d) : null;
    }
    const text = cleanText(value);
    if (!text) return null;
    const full = text.match(/(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/);
    if (full) return localDate(Number(full[1]), Number(full[2]), Number(full[3]));
    const chinese = text.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
    if (chinese) {
      return localDate(Number(chinese[1]), Number(chinese[2]), Number(chinese[3]));
    }
    const short = text.match(/^(\d{1,2})[\/.-](\d{1,2})$/);
    if (short && fallbackYear) {
      return localDate(fallbackYear, Number(short[1]), Number(short[2]));
    }
    return null;
  }

  function splitSpanText(text) {
    const normalized = cleanText(text).replace(/[～—–]/g, "~");
    const delimiters = ["~", "至"];
    for (const delimiter of delimiters) {
      const index = normalized.indexOf(delimiter);
      if (index > 0) {
        return [
          normalized.slice(0, index).trim(),
          normalized.slice(index + delimiter.length).trim(),
        ];
      }
    }
    return null;
  }

  function parseSpan(value, fallbackYear) {
    if (value instanceof Date || typeof value === "number") {
      const day = parseDate(value, fallbackYear);
      return day ? [day, day] : null;
    }
    const parts = splitSpanText(value);
    if (!parts) {
      const day = parseDate(value, fallbackYear);
      return day ? [day, day] : null;
    }
    const start = parseDate(parts[0], fallbackYear);
    const end = parseDate(parts[1], start ? start.getFullYear() : fallbackYear);
    return start && end ? [start, end] : null;
  }

  function addDays(date, days) {
    const result = new Date(date.getTime());
    result.setDate(result.getDate() + days);
    return result;
  }

  function maxDate(...dates) {
    return new Date(Math.max(...dates.map((value) => value.getTime())));
  }

  function minDate(...dates) {
    return new Date(Math.min(...dates.map((value) => value.getTime())));
  }

  function monthKey(value) {
    const iso = value instanceof Date ? dateToISO(value) : String(value);
    return iso.slice(0, 7);
  }

  function quarterFromDate(date) {
    const startMonth = Math.floor(date.getMonth() / 3) * 3 + 1;
    const start = localDate(date.getFullYear(), startMonth, 1);
    const end = new Date(date.getFullYear(), startMonth + 2, 0, 12, 0, 0, 0);
    const months = [0, 1, 2].map((offset) => {
      const day = localDate(date.getFullYear(), startMonth + offset, 1);
      return {
        key: monthKey(day),
        label: `${day.getMonth() + 1}月`,
        year: day.getFullYear(),
        month: day.getMonth() + 1,
      };
    });
    return {
      start,
      end,
      months,
      code: `${date.getFullYear()}Q${Math.floor(date.getMonth() / 3) + 1}`,
      shortCode: `Q${Math.floor(date.getMonth() / 3) + 1}`,
      label: `${date.getFullYear()}年${startMonth}-${startMonth + 2}月`,
    };
  }

  function endOfMonth(date) {
    return new Date(date.getFullYear(), date.getMonth() + 1, 0, 12, 0, 0, 0);
  }

  function isSameDay(left, right) {
    return dateToISO(left) === dateToISO(right);
  }

  function periodFromRange(startValue, endValue) {
    const start = localDate(
      startValue.getFullYear(),
      startValue.getMonth() + 1,
      startValue.getDate(),
    );
    const end = localDate(
      endValue.getFullYear(),
      endValue.getMonth() + 1,
      endValue.getDate(),
    );
    if (!start || !end || start > end) throw new Error("无法识别有效的统计日期范围");
    const months = [];
    for (
      let cursor = localDate(start.getFullYear(), start.getMonth() + 1, 1);
      cursor <= end;
      cursor = new Date(
        cursor.getFullYear(),
        cursor.getMonth() + 1,
        1,
        12,
        0,
        0,
        0,
      )
    ) {
      months.push({
        key: monthKey(cursor),
        year: cursor.getFullYear(),
        month: cursor.getMonth() + 1,
      });
    }
    const spansYears = months[0].year !== months[months.length - 1].year;
    months.forEach((month) => {
      month.label = spansYears ? `${month.year}年${month.month}月` : `${month.month}月`;
    });
    const fullMonthRange =
      start.getDate() === 1 && isSameDay(end, endOfMonth(end));
    let label;
    if (fullMonthRange && months.length === 1) {
      label = `${start.getFullYear()}年${start.getMonth() + 1}月`;
    } else if (fullMonthRange && !spansYears) {
      label = `${start.getFullYear()}年${start.getMonth() + 1}-${end.getMonth() + 1}月`;
    } else if (fullMonthRange) {
      label = `${start.getFullYear()}年${start.getMonth() + 1}月-${end.getFullYear()}年${
        end.getMonth() + 1
      }月`;
    } else if (
      start.getFullYear() === end.getFullYear() &&
      start.getMonth() === end.getMonth()
    ) {
      label = `${start.getFullYear()}年${start.getMonth() + 1}月${start.getDate()}-${
        end.getDate()
      }日`;
    } else {
      label = `${start.getFullYear()}年${start.getMonth() + 1}月${start.getDate()}日-${
        end.getFullYear()
      }年${end.getMonth() + 1}月${end.getDate()}日`;
    }
    const quarterStartMonth = Math.floor(start.getMonth() / 3) * 3 + 1;
    const quarterEnd = new Date(
      start.getFullYear(),
      quarterStartMonth + 2,
      0,
      12,
      0,
      0,
      0,
    );
    const isCalendarQuarter =
      fullMonthRange &&
      months.length === 3 &&
      start.getMonth() + 1 === quarterStartMonth &&
      isSameDay(end, quarterEnd);
    const quarterNumber = Math.floor(start.getMonth() / 3) + 1;
    return {
      start,
      end,
      months,
      label,
      code: isCalendarQuarter
        ? `${start.getFullYear()}Q${quarterNumber}`
        : `${dateToISO(start)}_${dateToISO(end)}`,
      shortCode: isCalendarQuarter ? `Q${quarterNumber}` : "统计期",
      isCalendarQuarter,
    };
  }

  function rowsFromSheet(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) throw new Error(`缺少工作表【${sheetName}】`);
    return getXlsx().utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: null,
      blankrows: false,
    });
  }

  function headerMap(row) {
    const result = new Map();
    (row || []).forEach((value, index) => {
      const header = cleanText(value);
      if (header) result.set(header, index);
    });
    return result;
  }

  function detectWorkbookYear(workbook) {
    const years = new Map();
    for (const sheetName of ["需求详情", "发单详情"]) {
      const rows = rowsFromSheet(workbook, sheetName);
      const headers = headerMap(rows[0]);
      const dateColumn = headers.get("需求日期");
      if (dateColumn == null) continue;
      for (let index = 1; index < rows.length; index += 1) {
        const raw = rows[index][dateColumn];
        const span = parseSpan(raw);
        if (!span) continue;
        for (const date of span) {
          years.set(date.getFullYear(), (years.get(date.getFullYear()) || 0) + 1);
        }
      }
    }
    const ranked = Array.from(years.entries()).sort((left, right) => {
      return right[1] - left[1] || left[0] - right[0];
    });
    return ranked.length ? ranked[0][0] : null;
  }

  function workbookPeriod(workbook, fallbackYear) {
    const rows = rowsFromSheet(workbook, "基本信息");
    if (!rows.length) throw new Error("【基本信息】为空");
    let year = detectWorkbookYear(workbook) || fallbackYear;
    if (!year) throw new Error("需求详情/发单详情未识别到有效年份");
    let previousMonth = null;
    const dates = [];
    for (let column = 3; column < (rows[0] || []).length; column += 1) {
      const raw = rows[0][column];
      if (raw instanceof Date || typeof raw === "number") {
        const parsed = parseDate(raw, year);
        if (parsed) dates.push(parsed);
        continue;
      }
      const match = cleanText(raw).match(/^(\d{1,2})-(\d{1,2})/);
      if (!match) continue;
      const month = Number(match[1]);
      if (previousMonth != null && month < previousMonth && previousMonth >= 11) {
        year += 1;
      }
      previousMonth = month;
      const parsed = localDate(year, month, Number(match[2]));
      if (parsed) dates.push(parsed);
    }
    if (!dates.length) throw new Error("【基本信息】未识别到日期列");
    return [new Date(Math.min(...dates)), new Date(Math.max(...dates))];
  }

  function parseBasic(source, workbook, period, analysisPeriod) {
    const rows = rowsFromSheet(workbook, "基本信息");
    const columnsByDate = new Map();
    let currentDay = null;
    for (let column = 3; column < Math.max(rows[0]?.length || 0, rows[1]?.length || 0); column += 1) {
      const header = rows[0]?.[column];
      const headerText = cleanText(header);
      if (header instanceof Date || typeof header === "number") {
        currentDay = parseDate(header, period[0].getFullYear());
      } else {
        const match = headerText.match(/^(\d{1,2})-(\d{1,2})/);
        if (match) {
          const candidates = [
            localDate(period[0].getFullYear(), Number(match[1]), Number(match[2])),
            localDate(period[1].getFullYear(), Number(match[1]), Number(match[2])),
          ].filter(Boolean);
          currentDay =
            candidates.find(
              (candidate) => candidate >= period[0] && candidate <= period[1],
            ) || candidates[0] || null;
        } else if (headerText) {
          currentDay = null;
        }
      }
      const metric = BASIC_METRICS[cleanText(rows[1]?.[column])];
      if (!currentDay || !metric) continue;
      const iso = dateToISO(currentDay);
      if (!columnsByDate.has(iso)) columnsByDate.set(iso, new Map());
      columnsByDate.get(iso).set(metric, column);
    }

    const result = [];
    for (let rowIndex = 2; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex];
      const region = cleanText(row[0]);
      if (!region) continue;
      const warehouse = cleanText(row[1]);
      const group = cleanText(row[2]);
      for (const [iso, metrics] of columnsByDate.entries()) {
        const day = parseDate(iso);
        if (
          !day ||
          day < period[0] ||
          day > period[1] ||
          day < analysisPeriod.start ||
          day > analysisPeriod.end
        ) {
          continue;
        }
        const record = {
          source_file: source.name,
          source_sheet: "基本信息",
          source_row: rowIndex + 1,
          region,
          warehouse,
          group,
          date: iso,
          month: monthKey(iso),
        };
        Object.values(BASIC_METRICS).forEach((metric) => {
          const column = metrics.get(metric);
          record[metric] = column == null ? 0 : asNumber(row[column]);
        });
        result.push(record);
      }
    }
    return result;
  }

  function requiredColumn(headers, name, sheetName) {
    const column = headers.get(name);
    if (column == null) throw new Error(`【${sheetName}】缺少字段“${name}”`);
    return column;
  }

  function parseDetail(source, workbook, period, analysisPeriod, sheetName) {
    const rows = rowsFromSheet(workbook, sheetName);
    const headers = headerMap(rows[0]);
    const isDispatch = sheetName === "发单详情";
    const columns = {
      region: requiredColumn(headers, "需求区域", sheetName),
      warehouse: requiredColumn(headers, "需求仓", sheetName),
      group: requiredColumn(headers, "需求组", sheetName),
      date: requiredColumn(headers, "需求日期", sheetName),
      job: requiredColumn(headers, "工种", sheetName),
      shift: requiredColumn(headers, "班次", sheetName),
      issued: requiredColumn(headers, "已发单人数", sheetName),
      pending_dispatch: requiredColumn(headers, "待派遣人数", sheetName),
      dispatched: requiredColumn(headers, "已派遣人数", sheetName),
    };
    if (isDispatch) {
      columns.supplier_name = requiredColumn(headers, "供应商名称", sheetName);
      columns.supplier_id = requiredColumn(headers, "供应商ID", sheetName);
      columns.order_type = headers.get("类型");
    } else {
      columns.demand = requiredColumn(headers, "需求人数", sheetName);
      columns.pending_issue = requiredColumn(headers, "待发单人数", sheetName);
      columns.application_id = headers.get("关联申请编号");
    }

    const expanded = [];
    const spans = [];
    const outside = [];
    const invalidDates = [];
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex];
      const region = cleanText(row[columns.region]);
      const rawDate = row[columns.date];
      if (!region || rawDate == null || rawDate === "") continue;
      const span = parseSpan(rawDate, period[0].getFullYear());
      if (!span) {
        invalidDates.push({
          source_file: source.name,
          sheet: sheetName,
          source_row: rowIndex + 1,
          raw_date: cleanText(rawDate),
        });
        continue;
      }
      const [start, end] = span;
      if (start.getTime() !== end.getTime()) {
        spans.push({
          source_file: source.name,
          sheet: sheetName,
          source_row: rowIndex + 1,
          raw_date: cleanText(rawDate),
          span_days: Math.round((end - start) / 86400000) + 1,
        });
      }
      const clippedStart = maxDate(start, period[0], analysisPeriod.start);
      const clippedEnd = minDate(end, period[1], analysisPeriod.end);
      if (clippedStart > clippedEnd) {
        outside.push({
          source_file: source.name,
          sheet: sheetName,
          source_row: rowIndex + 1,
          raw_date: cleanText(rawDate),
          file_period: `${dateToISO(period[0])}~${dateToISO(period[1])}`,
        });
        continue;
      }
      const base = {
        source_file: source.name,
        source_sheet: sheetName,
        source_row: rowIndex + 1,
        region,
        warehouse: cleanText(row[columns.warehouse]),
        group: cleanText(row[columns.group]),
        job: cleanText(row[columns.job]),
        shift: cleanText(row[columns.shift]),
        raw_date: cleanText(rawDate),
        issued: asNumber(row[columns.issued]),
        pending_dispatch: asNumber(row[columns.pending_dispatch]),
        dispatched: asNumber(row[columns.dispatched]),
      };
      if (isDispatch) {
        base.supplier_name = cleanText(row[columns.supplier_name]);
        base.supplier_id = cleanText(row[columns.supplier_id]);
        base.order_type = columns.order_type == null ? "" : cleanText(row[columns.order_type]);
      } else {
        base.demand = asNumber(row[columns.demand]);
        base.pending_issue = asNumber(row[columns.pending_issue]);
        base.application_id =
          columns.application_id == null ? "" : cleanText(row[columns.application_id]);
      }
      for (let day = clippedStart; day <= clippedEnd; day = addDays(day, 1)) {
        const record = { ...base, date: dateToISO(day), month: monthKey(day) };
        expanded.push(record);
      }
    }
    return { rows: expanded, spans, outside, invalidDates };
  }

  function keyOf(row, dimensions) {
    return dimensions.map((dimension) => row[dimension] ?? "").join("\u0001");
  }

  function splitKey(key) {
    return key.split("\u0001");
  }

  function aggregate(rows, dimensions, metrics) {
    const result = new Map();
    rows.forEach((row) => {
      const key = keyOf(row, dimensions);
      if (!result.has(key)) {
        result.set(key, Object.fromEntries(metrics.map((metric) => [metric, 0])));
      }
      const values = result.get(key);
      metrics.forEach((metric) => {
        values[metric] += asNumber(row[metric]);
      });
    });
    return result;
  }

  function cleanNumber(value, digits = 8) {
    const numeric = asNumber(value);
    if (Math.abs(numeric - Math.round(numeric)) <= EPS) return Math.round(numeric);
    return Number(numeric.toFixed(digits));
  }

  function safeRate(numerator, denominator) {
    return denominator > EPS ? numerator / denominator : null;
  }

  function monthAverage(values) {
    const valid = values.filter((value) => value != null && Number.isFinite(value));
    return valid.length ? valid.reduce((sum, value) => sum + value, 0) / valid.length : null;
  }

  function isTestSupplier(row) {
    return /测试|test|dummy|demo/i.test(row.supplier_name || "");
  }

  function compareSources(leftRows, rightRows, dimensions, metrics, leftName, rightName) {
    const left = aggregate(leftRows, dimensions, metrics);
    const right = aggregate(rightRows, dimensions, metrics);
    const keys = new Set([...left.keys(), ...right.keys()]);
    const mismatches = [];
    keys.forEach((key) => {
      const leftValues = left.get(key) || Object.fromEntries(metrics.map((metric) => [metric, 0]));
      const rightValues = right.get(key) || Object.fromEntries(metrics.map((metric) => [metric, 0]));
      const difference = Math.max(
        ...metrics.map((metric) => Math.abs(leftValues[metric] - rightValues[metric])),
      );
      if (difference <= EPS) return;
      const record = {};
      dimensions.forEach((dimension, index) => {
        record[dimension] = splitKey(key)[index];
      });
      metrics.forEach((metric) => {
        record[`${leftName}_${metric}`] = cleanNumber(leftValues[metric]);
        record[`${rightName}_${metric}`] = cleanNumber(rightValues[metric]);
      });
      record.max_abs_diff = cleanNumber(difference);
      mismatches.push(record);
    });
    return mismatches.sort((left, right) => right.max_abs_diff - left.max_abs_diff);
  }

  function cleanTestUnits(basicRows, demandRows, dispatchRows) {
    const testDispatchRows = dispatchRows.filter(isTestSupplier);
    const testAtomicKeys = new Set(testDispatchRows.map((row) => keyOf(row, DEMAND_DIMS)));
    const testGroupDays = new Set(
      testDispatchRows.map((row) => keyOf(row, DEMAND_DIMS.slice(0, 4))),
    );
    const cleanBasic = basicRows.filter(
      (row) => !testGroupDays.has(keyOf(row, DEMAND_DIMS.slice(0, 4))),
    );
    const cleanDemand = demandRows.filter(
      (row) => !testAtomicKeys.has(keyOf(row, DEMAND_DIMS)),
    );
    const cleanDispatch = dispatchRows.filter(
      (row) => !testAtomicKeys.has(keyOf(row, DEMAND_DIMS)) && !isTestSupplier(row),
    );

    const testUnits = Array.from(testAtomicKeys)
      .sort()
      .map((key) => {
        const values = splitKey(key);
        const record = Object.fromEntries(DEMAND_DIMS.map((dimension, index) => [dimension, values[index]]));
        const matchingDemand = demandRows.filter((row) => keyOf(row, DEMAND_DIMS) === key);
        const matchingDispatch = dispatchRows.filter((row) => keyOf(row, DEMAND_DIMS) === key);
        const groupKey = keyOf(record, DEMAND_DIMS.slice(0, 4));
        const matchingBasic = basicRows.filter(
          (row) => keyOf(row, DEMAND_DIMS.slice(0, 4)) === groupKey,
        );
        return {
          ...record,
          month: record.date.slice(0, 7),
          removed_basic_demand: cleanNumber(matchingBasic.reduce((sum, row) => sum + row.demand, 0)),
          removed_basic_supplement: cleanNumber(
            matchingBasic.reduce((sum, row) => sum + row.supplement, 0),
          ),
          removed_demand_detail: cleanNumber(
            matchingDemand.reduce((sum, row) => sum + row.demand, 0),
          ),
          removed_issued: cleanNumber(
            matchingDispatch.reduce((sum, row) => sum + row.issued, 0),
          ),
          removed_dispatched: cleanNumber(
            matchingDispatch.reduce((sum, row) => sum + row.dispatched, 0),
          ),
          test_suppliers: Array.from(
            new Set(
              matchingDispatch.map((row) => `${row.supplier_id} ${row.supplier_name}`.trim()),
            ),
          )
            .sort()
            .join(", "),
        };
      });
    return {
      basic_rows: cleanBasic,
      demand_rows: cleanDemand,
      dispatch_rows: cleanDispatch,
      test_units: testUnits,
      test_dispatch_rows: testDispatchRows,
    };
  }

  function buildAdjustedDemand(basicRows, demandRows) {
    const demand = aggregate(demandRows, DEMAND_DIMS, ["demand", "issued"]);
    const supplements = aggregate(
      basicRows,
      DEMAND_DIMS.slice(0, 4),
      ["supplement"],
    );
    const keysByGroupDay = new Map();
    demand.forEach((_values, key) => {
      const groupDay = splitKey(key).slice(0, 4).join("\u0001");
      if (!keysByGroupDay.has(groupDay)) keysByGroupDay.set(groupDay, []);
      keysByGroupDay.get(groupDay).push(key);
    });
    const adjusted = new Map();
    demand.forEach((values, key) => adjusted.set(key, values.demand));
    const allocations = [];

    supplements.forEach((values, groupDay) => {
      const supplement = values.supplement;
      if (supplement <= EPS) return;
      const candidates = keysByGroupDay.get(groupDay) || [];
      const groupValues = splitKey(groupDay);
      if (!candidates.length) {
        allocations.push({
          ...Object.fromEntries(
            DEMAND_DIMS.slice(0, 4).map((dimension, index) => [dimension, groupValues[index]]),
          ),
          supplement: cleanNumber(supplement),
          allocated_supplement: 0,
          method: "无法分配：需求详情无匹配行",
          candidate_count: 0,
        });
        return;
      }
      const demandTotal = candidates.reduce((sum, key) => sum + demand.get(key).demand, 0);
      const issuedTotal = candidates.reduce((sum, key) => sum + demand.get(key).issued, 0);
      let weights;
      let method;
      if (candidates.length === 1) {
        weights = [1];
        method = "唯一工种班次，全部归入";
      } else if (demandTotal > EPS) {
        weights = candidates.map((key) => demand.get(key).demand / demandTotal);
        method = "多工种班次，按原需求量占比分配";
      } else if (issuedTotal > EPS) {
        weights = candidates.map((key) => demand.get(key).issued / issuedTotal);
        method = "原需求为0，按发单量占比分配";
      } else {
        weights = candidates.map(() => 1 / candidates.length);
        method = "需求与发单均为0，平均分配";
      }
      candidates.forEach((key, index) => {
        const allocation = supplement * weights[index];
        adjusted.set(key, adjusted.get(key) + allocation);
        const values = splitKey(key);
        allocations.push({
          ...Object.fromEntries(
            DEMAND_DIMS.map((dimension, position) => [dimension, values[position]]),
          ),
          supplement: cleanNumber(supplement),
          allocated_supplement: cleanNumber(allocation),
          method,
          candidate_count: candidates.length,
        });
      });
    });

    const rows = [];
    demand.forEach((values, key) => {
      const parts = splitKey(key);
      const record = Object.fromEntries(
        DEMAND_DIMS.map((dimension, index) => [dimension, parts[index]]),
      );
      rows.push({
        ...record,
        month: record.date.slice(0, 7),
        raw_demand: cleanNumber(values.demand),
        adjusted_demand: cleanNumber(adjusted.get(key)),
      });
    });
    return { rows, allocations };
  }

  function buildAtomicTables(basicRows, demandRows, dispatchRows) {
    const adjusted = buildAdjustedDemand(basicRows, demandRows);
    const demandAtomic = aggregate(adjusted.rows, DEMAND_DIMS, [
      "raw_demand",
      "adjusted_demand",
    ]);
    const orderAtomic = aggregate(dispatchRows, DEMAND_DIMS, ["issued", "dispatched"]);
    const supplierBase = aggregate(dispatchRows, SUPPLIER_DIMS, ["issued", "dispatched"]);
    const keys = Array.from(new Set([...demandAtomic.keys(), ...orderAtomic.keys()])).sort();
    const unitRows = [];
    const unitLookup = new Map();
    keys.forEach((key) => {
      const parts = splitKey(key);
      const dimensions = Object.fromEntries(
        DEMAND_DIMS.map((dimension, index) => [dimension, parts[index]]),
      );
      const demand = demandAtomic.get(key) || { raw_demand: 0, adjusted_demand: 0 };
      const order = orderAtomic.get(key) || { issued: 0, dispatched: 0 };
      const rawDemand = demand.raw_demand;
      const trueDemand = demand.adjusted_demand;
      const issued = order.issued;
      const dispatched = order.dispatched;
      const factor = issued > EPS ? Math.min(1, trueDemand / issued) : 0;
      const row = {
        ...dimensions,
        month: dimensions.date.slice(0, 7),
        raw_demand: cleanNumber(rawDemand),
        supplement: cleanNumber(trueDemand - rawDemand),
        true_demand: cleanNumber(trueDemand),
        issued: cleanNumber(issued),
        dispatched: cleanNumber(dispatched),
        issue_covered: cleanNumber(Math.min(issued, trueDemand)),
        warehouse_effective: cleanNumber(Math.min(dispatched, trueDemand)),
        responsibility_factor: cleanNumber(factor),
        over_issue: cleanNumber(trueDemand > EPS ? Math.max(issued - trueDemand, 0) : 0),
        zero_demand_issue: cleanNumber(trueDemand <= EPS ? issued : 0),
        under_issue: cleanNumber(Math.max(trueDemand - issued, 0)),
        surplus_dispatch: cleanNumber(
          trueDemand > EPS ? Math.max(dispatched - trueDemand, 0) : dispatched,
        ),
      };
      unitRows.push(row);
      unitLookup.set(key, row);
    });

    const supplierRows = [];
    Array.from(supplierBase.keys())
      .sort()
      .forEach((key) => {
        const parts = splitKey(key);
        const dimensions = Object.fromEntries(
          SUPPLIER_DIMS.map((dimension, index) => [dimension, parts[index]]),
        );
        const unitKey = parts.slice(0, DEMAND_DIMS.length).join("\u0001");
        const unit = unitLookup.get(unitKey);
        if (!unit) throw new Error(`供应商原子行无仓库单元：${key}`);
        const values = supplierBase.get(key);
        const responsibility = values.issued * unit.responsibility_factor;
        supplierRows.push({
          ...dimensions,
          month: unit.month,
          true_demand: unit.true_demand,
          unit_total_issued: unit.issued,
          unit_total_dispatched: unit.dispatched,
          responsibility_factor: unit.responsibility_factor,
          issued: cleanNumber(values.issued),
          dispatched: cleanNumber(values.dispatched),
          raw_effective: cleanNumber(Math.min(values.dispatched, values.issued)),
          reasonable_responsibility: cleanNumber(responsibility),
          adjusted_effective: cleanNumber(Math.min(values.dispatched, responsibility)),
          over_issue_relief: cleanNumber(values.issued - responsibility),
          dispatch_over_order: cleanNumber(Math.max(values.dispatched - values.issued, 0)),
        });
      });
    return {
      unit_rows: unitRows,
      supplier_rows: supplierRows,
      supplement_allocations: adjusted.allocations,
    };
  }

  function aggregateToRecords(rows, dimensions, metrics) {
    const grouped = aggregate(rows, dimensions, metrics);
    return Array.from(grouped.entries()).map(([key, values]) => {
      const parts = splitKey(key);
      const record = Object.fromEntries(
        dimensions.map((dimension, index) => [dimension, parts[index]]),
      );
      metrics.forEach((metric) => {
        record[metric] = cleanNumber(values[metric]);
      });
      return record;
    });
  }

  function aggregateSupplierMonthly(supplierRows) {
    const dimensions = ["region", "supplier_id", "supplier_name", "job", "month"];
    const metrics = [
      "issued",
      "dispatched",
      "raw_effective",
      "reasonable_responsibility",
      "adjusted_effective",
      "over_issue_relief",
      "dispatch_over_order",
    ];
    return aggregateToRecords(supplierRows, dimensions, metrics)
      .map((row) => ({
        ...row,
        raw_rate: safeRate(row.raw_effective, row.issued),
        adjusted_rate: safeRate(row.adjusted_effective, row.reasonable_responsibility),
        rate_lift:
          row.issued > EPS && row.reasonable_responsibility > EPS
            ? row.adjusted_effective / row.reasonable_responsibility -
              row.raw_effective / row.issued
            : null,
      }))
      .sort((left, right) => {
        const leftKey = [
          left.region,
          left.supplier_id,
          left.supplier_name,
          left.job,
          left.month,
        ];
        const rightKey = [
          right.region,
          right.supplier_id,
          right.supplier_name,
          right.job,
          right.month,
        ];
        for (let index = 0; index < leftKey.length; index += 1) {
          if (leftKey[index] < rightKey[index]) return -1;
          if (leftKey[index] > rightKey[index]) return 1;
        }
        return 0;
      });
  }

  function aggregateWarehouseMonthly(unitRows) {
    const dimensions = ["region", "job", "month"];
    const metrics = [
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
    return aggregateToRecords(unitRows, dimensions, metrics)
      .map((row) => ({
        ...row,
        issue_coverage_rate: safeRate(row.issue_covered, row.true_demand),
        warehouse_fulfillment_rate: safeRate(row.warehouse_effective, row.true_demand),
      }))
      .sort((left, right) => {
        return (
          regionRank(left.region) - regionRank(right.region) ||
          left.job.localeCompare(right.job, "zh-CN") ||
          left.month.localeCompare(right.month)
        );
      });
  }

  function regionRank(region) {
    const index = REGION_ORDER.indexOf(region);
    if (index >= 0) return index;
    return REGION_ORDER.length + MARKET_ORDER.indexOf(marketForRegion(region));
  }

  function orderedRegions(rows) {
    const regions = Array.from(new Set(rows.map((row) => row.region)));
    return regions.sort((left, right) => {
      return regionRank(left) - regionRank(right) || left.localeCompare(right, "zh-CN");
    });
  }

  function buildDisplayRows(supplierMonthly, warehouseMonthly, geography) {
    const regions = geography.regions;
    const demandByJob = new Map();
    warehouseMonthly.forEach((row) => {
      demandByJob.set(row.job, (demandByJob.get(row.job) || 0) + row.true_demand);
    });
    const jobs = Array.from(demandByJob.keys()).sort((left, right) => {
      return demandByJob.get(right) - demandByJob.get(left) || left.localeCompare(right, "zh-CN");
    });
    const issuedBySupplier = new Map();
    const jobsBySupplier = new Map();
    supplierMonthly.forEach((row) => {
      const key = [row.region, row.supplier_id, row.supplier_name].join("\u0001");
      issuedBySupplier.set(key, (issuedBySupplier.get(key) || 0) + row.issued);
      if (!jobsBySupplier.has(key)) jobsBySupplier.set(key, new Set());
      jobsBySupplier.get(key).add(row.job);
    });
    const displayRows = [];
    regions.forEach((region) => {
      const market = geography.region_market[region];
      const suppliers = Array.from(issuedBySupplier.keys())
        .filter((key) => splitKey(key)[0] === region)
        .sort((left, right) => {
          return (
            issuedBySupplier.get(right) - issuedBySupplier.get(left) ||
            splitKey(left)[2].localeCompare(splitKey(right)[2], "zh-CN") ||
            splitKey(left)[1].localeCompare(splitKey(right)[1])
          );
        });
      suppliers.forEach((key) => {
        const [, supplierId, supplierName] = splitKey(key);
        jobs.forEach((job) => {
          if (!jobsBySupplier.get(key).has(job)) return;
          displayRows.push({
            market,
            region,
            supplier_name: supplierName,
            supplier_id: supplierId,
            level: "工种明细",
            job,
            row_type: "supplier_job",
          });
        });
        displayRows.push({
          market,
          region,
          supplier_name: supplierName,
          supplier_id: supplierId,
          level: "供应商汇总",
          job: "全部工种",
          row_type: "supplier_total",
        });
      });
      const regionJobs = new Set(
        supplierMonthly.filter((row) => row.region === region).map((row) => row.job),
      );
      jobs.forEach((job) => {
        if (!regionJobs.has(job)) return;
        displayRows.push({
          market,
          region,
          supplier_name: "全区",
          supplier_id: "",
          level: "区域汇总",
          job,
          row_type: "region_job",
        });
      });
      displayRows.push({
        market,
        region,
        supplier_name: "全区",
        supplier_id: "",
        level: "区域汇总",
        job: "全部工种",
        row_type: "region_total",
      });
    });

    if (geography.markets.length > 1) {
      geography.markets.forEach((market) => {
        const memberRegions = regions.filter(
          (region) => geography.region_market[region] === market,
        );
        const marketJobs = new Set(
          supplierMonthly
            .filter((row) => memberRegions.includes(row.region))
            .map((row) => row.job),
        );
        jobs.forEach((job) => {
          if (!marketJobs.has(job)) return;
          displayRows.push({
            market,
            region: `${market}总计`,
            supplier_name: `${market}汇总`,
            supplier_id: "",
            level: `${market}汇总`,
            job,
            row_type: "market_job",
            member_regions: memberRegions,
          });
        });
        displayRows.push({
          market,
          region: `${market}总计`,
          supplier_name: `${market}汇总`,
          supplier_id: "",
          level: `${market}汇总`,
          job: "全部工种",
          row_type: "market_total",
          member_regions: memberRegions,
        });
      });
    }

    if (regions.length > 1) {
      jobs.forEach((job) => {
        if (!supplierMonthly.some((row) => row.job === job)) return;
        displayRows.push({
          market: geography.prefix,
          region: geography.scope_total_label,
          supplier_name: geography.scope_supplier_name,
          supplier_id: "",
          level: geography.scope_level,
          job,
          row_type: "scope_job",
        });
      });
      displayRows.push({
        market: geography.prefix,
        region: geography.scope_total_label,
        supplier_name: geography.scope_supplier_name,
        supplier_id: "",
        level: geography.scope_level,
        job: "全部工种",
        row_type: "scope_total",
      });
    }
    return { displayRows, regions, jobs, geography };
  }

  function rowMatchesDisplay(monthlyRow, displayRow) {
    if (
      ["supplier_job", "supplier_total", "region_job", "region_total"].includes(
        displayRow.row_type,
      ) &&
      monthlyRow.region !== displayRow.region
    ) {
      return false;
    }
    if (
      ["market_job", "market_total"].includes(displayRow.row_type) &&
      !displayRow.member_regions.includes(monthlyRow.region)
    ) {
      return false;
    }
    if (
      ["supplier_job", "supplier_total"].includes(displayRow.row_type) &&
      monthlyRow.supplier_id !== displayRow.supplier_id
    ) {
      return false;
    }
    if (
      ["supplier_job", "region_job", "market_job", "scope_job"].includes(
        displayRow.row_type,
      ) &&
      monthlyRow.job !== displayRow.job
    ) {
      return false;
    }
    return true;
  }

  function buildDisplayMetrics(displayRows, supplierMonthly, months) {
    return displayRows.map((displayRow) => {
      const matching = supplierMonthly.filter((row) => rowMatchesDisplay(row, displayRow));
      const byMonth = months.map((month) => {
        const rows = matching.filter((row) => row.month === month.key);
        const issued = rows.reduce((sum, row) => sum + row.issued, 0);
        const rawEffective = rows.reduce((sum, row) => sum + row.raw_effective, 0);
        const responsibility = rows.reduce(
          (sum, row) => sum + row.reasonable_responsibility,
          0,
        );
        const adjustedEffective = rows.reduce(
          (sum, row) => sum + row.adjusted_effective,
          0,
        );
        const relief = rows.reduce((sum, row) => sum + row.over_issue_relief, 0);
        return {
          month: month.key,
          issued: cleanNumber(issued),
          raw_effective: cleanNumber(rawEffective),
          raw_rate: safeRate(rawEffective, issued),
          reasonable_responsibility: cleanNumber(responsibility),
          adjusted_effective: cleanNumber(adjustedEffective),
          adjusted_rate: safeRate(adjustedEffective, responsibility),
          over_issue_relief: cleanNumber(relief),
        };
      });
      const totals = byMonth.reduce(
        (result, row) => {
          result.issued += row.issued;
          result.raw_effective += row.raw_effective;
          result.reasonable_responsibility += row.reasonable_responsibility;
          result.adjusted_effective += row.adjusted_effective;
          result.over_issue_relief += row.over_issue_relief;
          return result;
        },
        {
          issued: 0,
          raw_effective: 0,
          reasonable_responsibility: 0,
          adjusted_effective: 0,
          over_issue_relief: 0,
        },
      );
      return {
        ...displayRow,
        by_month: byMonth,
        raw_average: monthAverage(byMonth.map((row) => row.raw_rate)),
        adjusted_average: monthAverage(byMonth.map((row) => row.adjusted_rate)),
        raw_weighted: safeRate(totals.raw_effective, totals.issued),
        adjusted_weighted: safeRate(
          totals.adjusted_effective,
          totals.reasonable_responsibility,
        ),
        totals: Object.fromEntries(
          Object.entries(totals).map(([key, value]) => [key, cleanNumber(value)]),
        ),
      };
    });
  }

  function buildWarehouseDisplay(
    warehouseMonthly,
    supplierMonthly,
    regions,
    jobs,
    months,
    geography,
  ) {
    const observed = new Set(
      warehouseMonthly.map((row) => [row.region, row.job, row.month].join("\u0001")),
    );
    const regionMonths = new Set(
      warehouseMonthly.map((row) => [row.region, row.month].join("\u0001")),
    );
    const jobMonths = new Set(
      warehouseMonthly.map((row) => [row.job, row.month].join("\u0001")),
    );
    const allMonths = new Set(warehouseMonthly.map((row) => row.month));
    const rows = [];
    regions.forEach((region) => {
      const market = geography.region_market[region];
      const regionJobs = new Set(
        warehouseMonthly.filter((row) => row.region === region).map((row) => row.job),
      );
      jobs.forEach((job) => {
        if (!regionJobs.has(job)) return;
        months.forEach((month) => {
          if (!observed.has([region, job, month.key].join("\u0001"))) return;
          rows.push({
            market,
            region,
            level: "工种明细",
            job,
            month: month.key,
            row_type: "region_job",
          });
        });
      });
      months.forEach((month) => {
        if (!regionMonths.has([region, month.key].join("\u0001"))) return;
        rows.push({
          market,
          region,
          level: "区域汇总",
          job: "全部工种",
          month: month.key,
          row_type: "region_total",
        });
      });
    });
    if (geography.markets.length > 1) {
      geography.markets.forEach((market) => {
        const memberRegions = regions.filter(
          (region) => geography.region_market[region] === market,
        );
        jobs.forEach((job) => {
          months.forEach((month) => {
            if (
              !warehouseMonthly.some(
                (row) =>
                  memberRegions.includes(row.region) &&
                  row.job === job &&
                  row.month === month.key,
              )
            ) {
              return;
            }
            rows.push({
              market,
              region: `${market}总计`,
              level: `${market}汇总`,
              job,
              month: month.key,
              row_type: "market_job",
              member_regions: memberRegions,
            });
          });
        });
        months.forEach((month) => {
          if (
            !warehouseMonthly.some(
              (row) => memberRegions.includes(row.region) && row.month === month.key,
            )
          ) {
            return;
          }
          rows.push({
            market,
            region: `${market}总计`,
            level: `${market}汇总`,
            job: "全部工种",
            month: month.key,
            row_type: "market_total",
            member_regions: memberRegions,
          });
        });
      });
    }
    if (regions.length > 1) {
      jobs.forEach((job) => {
        months.forEach((month) => {
          if (!jobMonths.has([job, month.key].join("\u0001"))) return;
          rows.push({
            market: geography.prefix,
            region: geography.scope_total_label,
            level: geography.scope_level,
            job,
            month: month.key,
            row_type: "scope_job",
          });
        });
      });
      months.forEach((month) => {
        if (!allMonths.has(month.key)) return;
        rows.push({
          market: geography.prefix,
          region: geography.scope_total_label,
          level: geography.scope_level,
          job: "全部工种",
          month: month.key,
          row_type: "scope_total",
        });
      });
    }

    function matches(row, display) {
      if (
        ["region_job", "region_total"].includes(display.row_type) &&
        row.region !== display.region
      ) {
        return false;
      }
      if (
        ["market_job", "market_total"].includes(display.row_type) &&
        !display.member_regions.includes(row.region)
      ) {
        return false;
      }
      if (
        ["region_job", "market_job", "scope_job"].includes(display.row_type) &&
        row.job !== display.job
      ) {
        return false;
      }
      return row.month === display.month;
    }

    return rows.map((display) => {
      const warehouseRows = warehouseMonthly.filter((row) => matches(row, display));
      const supplierRows = supplierMonthly.filter((row) => matches(row, display));
      const metrics = [
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
      const totals = Object.fromEntries(
        metrics.map((metric) => [
          metric,
          cleanNumber(warehouseRows.reduce((sum, row) => sum + row[metric], 0)),
        ]),
      );
      const rawDen = supplierRows.reduce((sum, row) => sum + row.issued, 0);
      const rawNum = supplierRows.reduce((sum, row) => sum + row.raw_effective, 0);
      const adjustedDen = supplierRows.reduce(
        (sum, row) => sum + row.reasonable_responsibility,
        0,
      );
      const adjustedNum = supplierRows.reduce(
        (sum, row) => sum + row.adjusted_effective,
        0,
      );
      return {
        ...display,
        ...totals,
        issue_coverage_rate: safeRate(totals.issue_covered, totals.true_demand),
        warehouse_fulfillment_rate: safeRate(
          totals.warehouse_effective,
          totals.true_demand,
        ),
        raw_rate: safeRate(rawNum, rawDen),
        adjusted_rate: safeRate(adjustedNum, adjustedDen),
      };
    });
  }

  function sourceMonthTotals(basicRows, demandRows, dispatchRows, months) {
    const basic = aggregate(basicRows, ["month"], [
      "demand",
      "supplement",
      "issued",
      "dispatched",
    ]);
    const demand = aggregate(demandRows, ["month"], ["demand", "issued", "dispatched"]);
    const dispatch = aggregate(dispatchRows, ["month"], ["issued", "dispatched"]);
    return months.map((month) => {
      const key = month.key;
      const b = basic.get(key) || {};
      const d = demand.get(key) || {};
      const o = dispatch.get(key) || {};
      return {
        month: key,
        basic_demand: cleanNumber(b.demand),
        basic_supplement: cleanNumber(b.supplement),
        basic_issued: cleanNumber(b.issued),
        basic_dispatched: cleanNumber(b.dispatched),
        demand_detail_demand: cleanNumber(d.demand),
        demand_detail_issued: cleanNumber(d.issued),
        demand_detail_dispatched: cleanNumber(d.dispatched),
        dispatch_detail_issued: cleanNumber(o.issued),
        dispatch_detail_dispatched: cleanNumber(o.dispatched),
      };
    });
  }

  function monthRegionReconciliation(basicRows, demandRows, dispatchRows) {
    const dimensions = ["month", "region"];
    const basic = aggregate(basicRows, dimensions, ["demand", "issued", "dispatched"]);
    const demand = aggregate(demandRows, dimensions, ["demand", "issued", "dispatched"]);
    const dispatch = aggregate(dispatchRows, dimensions, ["issued", "dispatched"]);
    const keys = Array.from(new Set([...basic.keys(), ...demand.keys(), ...dispatch.keys()]));
    return keys
      .map((key) => {
        const [month, region] = splitKey(key);
        const b = basic.get(key) || {};
        const d = demand.get(key) || {};
        const o = dispatch.get(key) || {};
        const values = {
          month,
          region,
          basic_demand: cleanNumber(b.demand),
          demand_detail_demand: cleanNumber(d.demand),
          basic_issued: cleanNumber(b.issued),
          demand_detail_issued: cleanNumber(d.issued),
          dispatch_detail_issued: cleanNumber(o.issued),
          basic_dispatched: cleanNumber(b.dispatched),
          demand_detail_dispatched: cleanNumber(d.dispatched),
          dispatch_detail_dispatched: cleanNumber(o.dispatched),
        };
        const issueValues = [
          values.basic_issued,
          values.demand_detail_issued,
          values.dispatch_detail_issued,
        ];
        const dispatchValues = [
          values.basic_dispatched,
          values.demand_detail_dispatched,
          values.dispatch_detail_dispatched,
        ];
        values.demand_difference = cleanNumber(
          values.basic_demand - values.demand_detail_demand,
        );
        values.issue_max_difference = cleanNumber(
          Math.max(...issueValues) - Math.min(...issueValues),
        );
        values.dispatch_max_difference = cleanNumber(
          Math.max(...dispatchValues) - Math.min(...dispatchValues),
        );
        values.max_abs_difference = cleanNumber(
          Math.max(
            Math.abs(values.demand_difference),
            values.issue_max_difference,
            values.dispatch_max_difference,
          ),
        );
        values.status = values.max_abs_difference <= EPS ? "一致" : "不一致";
        return values;
      })
      .sort((left, right) => {
        return left.month.localeCompare(right.month) || regionRank(left.region) - regionRank(right.region);
      });
  }

  function buildQuality(raw, cleaned, atomic, supplierMonthly, reconciliation) {
    const responsibilityByUnit = aggregate(
      atomic.supplier_rows,
      DEMAND_DIMS,
      ["reasonable_responsibility"],
    );
    const responsibilityClosure = atomic.unit_rows
      .map((unit) => {
        const key = keyOf(unit, DEMAND_DIMS);
        const expected = Math.min(unit.issued, unit.true_demand);
        const actual =
          responsibilityByUnit.get(key)?.reasonable_responsibility || 0;
        return {
          ...Object.fromEntries(DEMAND_DIMS.map((dimension) => [dimension, unit[dimension]])),
          expected: cleanNumber(expected),
          actual: cleanNumber(actual),
          difference: cleanNumber(actual - expected),
        };
      })
      .filter((row) => Math.abs(row.difference) > 1e-6);
    const demandVsDispatch = compareSources(
      cleaned.demand_rows,
      cleaned.dispatch_rows,
      DEMAND_DIMS,
      ["issued", "dispatched"],
      "demand_detail",
      "dispatch_detail",
    );
    const basicVsDemand = compareSources(
      cleaned.basic_rows,
      cleaned.demand_rows,
      DEMAND_DIMS.slice(0, 4),
      ["demand", "pending_issue", "issued", "pending_dispatch", "dispatched"],
      "basic",
      "demand_detail",
    );
    const positiveDemandNoIssue = atomic.unit_rows.filter(
      (row) => row.true_demand > EPS && row.issued <= EPS,
    );
    const zeroDemandWithIssue = atomic.unit_rows.filter(
      (row) => row.true_demand <= EPS && row.issued > EPS,
    );
    const overIssue = atomic.unit_rows.filter(
      (row) => row.true_demand > EPS && row.issued > row.true_demand + EPS,
    );
    const overDispatch = atomic.unit_rows.filter(
      (row) => row.dispatched > row.true_demand + EPS,
    );
    const supplierDispatchOverOrder = atomic.supplier_rows.filter(
      (row) => row.dispatched > row.issued + EPS,
    );
    const invalidRates = supplierMonthly.filter((row) => {
      return [row.raw_rate, row.adjusted_rate].some(
        (rate) => rate != null && (rate < -EPS || rate > 1 + EPS),
      );
    });
    const adjustedBelowRaw = supplierMonthly.filter(
      (row) =>
        row.raw_rate != null &&
        row.adjusted_rate != null &&
        row.adjusted_rate + EPS < row.raw_rate,
    );
    const rawSupplierIssued = atomic.supplier_rows.reduce((sum, row) => sum + row.issued, 0);
    const rawUnitIssued = atomic.unit_rows.reduce((sum, row) => sum + row.issued, 0);
    const rawSupplierDispatched = atomic.supplier_rows.reduce(
      (sum, row) => sum + row.dispatched,
      0,
    );
    const rawUnitDispatched = atomic.unit_rows.reduce(
      (sum, row) => sum + row.dispatched,
      0,
    );
    const criticalChecks = [
      {
        id: "three_way_month_region",
        label: "三表月×区域对账",
        status: reconciliation.every((row) => row.max_abs_difference <= EPS)
          ? "PASS"
          : "FAIL",
        difference_count: reconciliation.filter((row) => row.max_abs_difference > EPS).length,
      },
      {
        id: "demand_dispatch_atomic",
        label: "需求详情与发单详情原子对账",
        status: demandVsDispatch.length === 0 ? "PASS" : "FAIL",
        difference_count: demandVsDispatch.length,
      },
      {
        id: "responsibility_closure",
        label: "合理责任闭合",
        status: responsibilityClosure.length === 0 ? "PASS" : "FAIL",
        difference_count: responsibilityClosure.length,
      },
      {
        id: "issued_closure",
        label: "供应商发单与仓库总发单闭合",
        status: Math.abs(rawSupplierIssued - rawUnitIssued) <= 1e-6 ? "PASS" : "FAIL",
        difference_count: Math.abs(rawSupplierIssued - rawUnitIssued) <= 1e-6 ? 0 : 1,
      },
      {
        id: "dispatch_closure",
        label: "供应商派遣与仓库总派遣闭合",
        status:
          Math.abs(rawSupplierDispatched - rawUnitDispatched) <= 1e-6 ? "PASS" : "FAIL",
        difference_count:
          Math.abs(rawSupplierDispatched - rawUnitDispatched) <= 1e-6 ? 0 : 1,
      },
      {
        id: "rate_bounds",
        label: "履约率边界",
        status: invalidRates.length === 0 ? "PASS" : "FAIL",
        difference_count: invalidRates.length,
      },
    ];
    return {
      source_row_counts: {
        basic_expanded: raw.basic_rows.length,
        demand_expanded: raw.demand_rows.length,
        dispatch_expanded: raw.dispatch_rows.length,
        date_range_source_rows: raw.span_rows.length,
        outside_period_rows: raw.outside_rows.length,
        invalid_date_rows: raw.invalid_date_rows.length,
      },
      clean_row_counts: {
        basic_expanded: cleaned.basic_rows.length,
        demand_expanded: cleaned.demand_rows.length,
        dispatch_expanded: cleaned.dispatch_rows.length,
        unit_atomic: atomic.unit_rows.length,
        supplier_atomic: atomic.supplier_rows.length,
        supplier_monthly: supplierMonthly.length,
      },
      month_region_reconciliation: reconciliation,
      month_region_reconciliation_mismatch_count: reconciliation.filter(
        (row) => row.max_abs_difference > EPS,
      ).length,
      basic_vs_demand_mismatches: basicVsDemand,
      basic_vs_demand_mismatch_count: basicVsDemand.length,
      demand_vs_dispatch_mismatches: demandVsDispatch,
      demand_vs_dispatch_mismatch_count: demandVsDispatch.length,
      responsibility_closure_mismatches: responsibilityClosure,
      responsibility_closure_mismatch_count: responsibilityClosure.length,
      invalid_supplier_rate_count: invalidRates.length,
      adjusted_below_raw_count: adjustedBelowRaw.length,
      adjusted_below_raw_rows: adjustedBelowRaw,
      supplier_dispatch_over_order_count: supplierDispatchOverOrder.length,
      positive_demand_no_issue: {
        unit_count: positiveDemandNoIssue.length,
        demand_person_days: cleanNumber(
          positiveDemandNoIssue.reduce((sum, row) => sum + row.true_demand, 0),
        ),
      },
      zero_demand_with_issue: {
        unit_count: zeroDemandWithIssue.length,
        issued_person_days: cleanNumber(
          zeroDemandWithIssue.reduce((sum, row) => sum + row.issued, 0),
        ),
        dispatched_person_days: cleanNumber(
          zeroDemandWithIssue.reduce((sum, row) => sum + row.dispatched, 0),
        ),
      },
      over_issue: {
        unit_count: overIssue.length,
        person_days: cleanNumber(overIssue.reduce((sum, row) => sum + row.over_issue, 0)),
      },
      over_dispatch: {
        unit_count: overDispatch.length,
        person_days: cleanNumber(
          overDispatch.reduce((sum, row) => sum + row.surplus_dispatch, 0),
        ),
      },
      test_unit_count: cleaned.test_units.length,
      test_units: cleaned.test_units,
      supplement_allocations: atomic.supplement_allocations,
      critical_checks: criticalChecks,
      status: criticalChecks.every((check) => check.status === "PASS") ? "PASS" : "REVIEW",
    };
  }

  function scopeNames(geography, period) {
    const prefix = geography.prefix;
    return {
      prefix,
      report_name: `${prefix}供应商履约率报表`,
      title: `${prefix}供应商履约率报表（${period.label}）`,
      filename: `${prefix}供应商履约率报表_${period.label}_最终严谨版.xlsx`,
      management_title: `${prefix}供应商履约率管理摘要（${period.label}）`,
    };
  }

  function buildScopeSummary(
    displayMetrics,
    warehouseDisplay,
    unitRows,
    supplierRows,
    regions,
    months,
  ) {
    const scopeRow =
      displayMetrics.find((row) => row.row_type === "scope_total") ||
      displayMetrics.find(
        (row) => row.row_type === "region_total" && row.region === regions[0],
      );
    const monthly = months.map((month) => {
      const displayMonth = scopeRow.by_month.find((row) => row.month === month.key);
      const warehouseRows = warehouseDisplay.filter((row) => {
        if (regions.length === 1) {
          return (
            row.row_type === "region_total" &&
            row.region === regions[0] &&
            row.month === month.key
          );
        }
        return row.row_type === "scope_total" && row.month === month.key;
      });
      const warehouse = warehouseRows[0] || {};
      return {
        month: month.key,
        label: month.label,
        raw_rate: displayMonth?.raw_rate ?? null,
        adjusted_rate: displayMonth?.adjusted_rate ?? null,
        warehouse_rate: warehouse.warehouse_fulfillment_rate ?? null,
        issue_coverage_rate: warehouse.issue_coverage_rate ?? null,
        issue_to_demand: safeRate(warehouse.issued || 0, warehouse.true_demand || 0),
        dispatch_to_demand: safeRate(
          warehouse.dispatched || 0,
          warehouse.true_demand || 0,
        ),
      };
    });
    const trueDemand = unitRows.reduce((sum, row) => sum + row.true_demand, 0);
    const issued = unitRows.reduce((sum, row) => sum + row.issued, 0);
    const dispatched = unitRows.reduce((sum, row) => sum + row.dispatched, 0);
    const warehouseEffective = unitRows.reduce(
      (sum, row) => sum + row.warehouse_effective,
      0,
    );
    const issueCovered = unitRows.reduce((sum, row) => sum + row.issue_covered, 0);
    const responsibility = supplierRows.reduce(
      (sum, row) => sum + row.reasonable_responsibility,
      0,
    );
    const adjustedEffective = supplierRows.reduce(
      (sum, row) => sum + row.adjusted_effective,
      0,
    );
    const rawEffective = supplierRows.reduce((sum, row) => sum + row.raw_effective, 0);
    return {
      monthly,
      averages: {
        raw_rate: monthAverage(monthly.map((row) => row.raw_rate)),
        adjusted_rate: monthAverage(monthly.map((row) => row.adjusted_rate)),
        warehouse_rate: monthAverage(monthly.map((row) => row.warehouse_rate)),
        issue_coverage_rate: monthAverage(monthly.map((row) => row.issue_coverage_rate)),
        issue_to_demand: monthAverage(monthly.map((row) => row.issue_to_demand)),
        dispatch_to_demand: monthAverage(monthly.map((row) => row.dispatch_to_demand)),
      },
      weighted: {
        raw_rate: safeRate(rawEffective, issued),
        adjusted_rate: safeRate(adjustedEffective, responsibility),
        warehouse_rate: safeRate(warehouseEffective, trueDemand),
        issue_coverage_rate: safeRate(issueCovered, trueDemand),
        issue_to_demand: safeRate(issued, trueDemand),
        dispatch_to_demand: safeRate(dispatched, trueDemand),
      },
    };
  }

  function buildRegionSummary(displayMetrics, warehouseDisplay, unitRows, regions) {
    return regions.map((region) => {
      const row = displayMetrics.find(
        (item) => item.row_type === "region_total" && item.region === region,
      );
      const warehouseRows = warehouseDisplay.filter(
        (item) => item.row_type === "region_total" && item.region === region,
      );
      return {
        market: marketForRegion(region),
        region,
        label: region,
        row_type: "region_total",
        member_regions: [region],
        raw_average: row?.raw_average ?? null,
        adjusted_average: row?.adjusted_average ?? null,
        average_lift:
          row?.raw_average != null && row?.adjusted_average != null
            ? row.adjusted_average - row.raw_average
            : null,
        warehouse_average: monthAverage(
          warehouseRows.map((item) => item.warehouse_fulfillment_rate),
        ),
        coverage_average: monthAverage(
          warehouseRows.map((item) => item.issue_coverage_rate),
        ),
        over_issue: cleanNumber(
          unitRows
            .filter((item) => item.region === region)
            .reduce((sum, item) => sum + item.over_issue, 0),
        ),
        under_issue: cleanNumber(
          unitRows
            .filter((item) => item.region === region)
            .reduce((sum, item) => sum + item.under_issue, 0),
        ),
      };
    });
  }

  function buildMarketSummary(
    displayMetrics,
    warehouseDisplay,
    unitRows,
    geography,
  ) {
    if (geography.markets.length <= 1) return [];
    return geography.markets.map((market) => {
      const label = `${market}总计`;
      const memberRegions = geography.regions.filter(
        (region) => geography.region_market[region] === market,
      );
      const row = displayMetrics.find(
        (item) => item.row_type === "market_total" && item.market === market,
      );
      const warehouseRows = warehouseDisplay.filter(
        (item) => item.row_type === "market_total" && item.market === market,
      );
      return {
        market,
        region: label,
        label,
        row_type: "market_total",
        member_regions: memberRegions,
        raw_average: row?.raw_average ?? null,
        adjusted_average: row?.adjusted_average ?? null,
        average_lift:
          row?.raw_average != null && row?.adjusted_average != null
            ? row.adjusted_average - row.raw_average
            : null,
        warehouse_average: monthAverage(
          warehouseRows.map((item) => item.warehouse_fulfillment_rate),
        ),
        coverage_average: monthAverage(
          warehouseRows.map((item) => item.issue_coverage_rate),
        ),
        over_issue: cleanNumber(
          unitRows
            .filter((item) => memberRegions.includes(item.region))
            .reduce((sum, item) => sum + item.over_issue, 0),
        ),
        under_issue: cleanNumber(
          unitRows
            .filter((item) => memberRegions.includes(item.region))
            .reduce((sum, item) => sum + item.under_issue, 0),
        ),
      };
    });
  }

  function coverageAudit(periods, analysisPeriod) {
    const covered = new Set();
    periods.forEach(([start, end]) => {
      const clippedStart = maxDate(start, analysisPeriod.start);
      const clippedEnd = minDate(end, analysisPeriod.end);
      for (let day = clippedStart; day <= clippedEnd; day = addDays(day, 1)) {
        covered.add(dateToISO(day));
      }
    });
    const expected =
      Math.round((analysisPeriod.end - analysisPeriod.start) / 86400000) + 1;
    const missing = [];
    for (
      let day = analysisPeriod.start;
      day <= analysisPeriod.end;
      day = addDays(day, 1)
    ) {
      const iso = dateToISO(day);
      if (!covered.has(iso)) missing.push(iso);
    }
    return {
      expected_days: expected,
      covered_days: covered.size,
      missing_days: missing,
      complete: covered.size === expected,
    };
  }

  async function sha256(arrayBuffer) {
    if (root.crypto?.subtle) {
      const digest = await root.crypto.subtle.digest("SHA-256", arrayBuffer.slice(0));
      return Array.from(new Uint8Array(digest))
        .map((value) => value.toString(16).padStart(2, "0"))
        .join("");
    }
    if (typeof require === "function") {
      const crypto = require("crypto");
      return crypto.createHash("sha256").update(Buffer.from(arrayBuffer)).digest("hex");
    }
    return "浏览器不支持SHA-256";
  }

  async function normalizeSources(files) {
    const result = [];
    for (const file of Array.from(files || [])) {
      const arrayBuffer =
        file.arrayBuffer instanceof Function ? await file.arrayBuffer() : file.data;
      if (!(arrayBuffer instanceof ArrayBuffer)) {
        throw new Error(`${file.name || "未命名文件"}无法读取为Excel二进制`);
      }
      result.push({
        name: file.name || "未命名文件.xlsx",
        size: file.size ?? arrayBuffer.byteLength,
        lastModified: file.lastModified || 0,
        arrayBuffer,
      });
    }
    return result;
  }

  function fileModifiedISO(source) {
    return source.lastModified
      ? new Date(source.lastModified).toISOString().slice(0, 19)
      : "";
  }

  async function analyzeFiles(files) {
    const sources = await normalizeSources(files);
    if (!sources.length) throw new Error("请至少上传一个用工需求池Excel文件");
    const loadedFiles = sources.map((source) => {
      let workbook;
      try {
        workbook = getXlsx().read(source.arrayBuffer, {
          type: "array",
          cellDates: true,
          raw: true,
          dense: false,
        });
      } catch (error) {
        throw new Error(`${source.name}无法解析：${error.message}`);
      }
      ["基本信息", "需求详情", "发单详情"].forEach((sheetName) => {
        if (!workbook.SheetNames.includes(sheetName)) {
          throw new Error(`${source.name}缺少工作表【${sheetName}】`);
        }
      });
      return { source, workbook };
    });
    const detectedYears = loadedFiles
      .map((item) => detectWorkbookYear(item.workbook))
      .filter((year) => year != null);
    if (!detectedYears.length) {
      throw new Error("全部文件的需求详情/发单详情均未识别到有效年份");
    }
    const yearCounts = new Map();
    detectedYears.forEach((year) => {
      yearCounts.set(year, (yearCounts.get(year) || 0) + 1);
    });
    const fallbackYear = Array.from(yearCounts.entries()).sort(
      (left, right) => right[1] - left[1] || left[0] - right[0],
    )[0][0];
    const parsedFiles = loadedFiles.map((item) => ({
      ...item,
      period: workbookPeriod(item.workbook, fallbackYear),
    }));
    const analysisPeriod = periodFromRange(
      parsedFiles.reduce(
        (earliest, item) => (item.period[0] < earliest ? item.period[0] : earliest),
        parsedFiles[0].period[0],
      ),
      parsedFiles.reduce(
        (latest, item) => (item.period[1] > latest ? item.period[1] : latest),
        parsedFiles[0].period[1],
      ),
    );

    const raw = {
      basic_rows: [],
      demand_rows: [],
      dispatch_rows: [],
      span_rows: [],
      outside_rows: [],
      invalid_date_rows: [],
      file_meta: [],
    };
    for (const item of parsedFiles) {
      const basicRows = parseBasic(
        item.source,
        item.workbook,
        item.period,
        analysisPeriod,
      );
      const demand = parseDetail(
        item.source,
        item.workbook,
        item.period,
        analysisPeriod,
        "需求详情",
      );
      const dispatch = parseDetail(
        item.source,
        item.workbook,
        item.period,
        analysisPeriod,
        "发单详情",
      );
      raw.basic_rows.push(...basicRows);
      raw.demand_rows.push(...demand.rows);
      raw.dispatch_rows.push(...dispatch.rows);
      raw.span_rows.push(...demand.spans, ...dispatch.spans);
      raw.outside_rows.push(...demand.outside, ...dispatch.outside);
      raw.invalid_date_rows.push(...demand.invalidDates, ...dispatch.invalidDates);
      raw.file_meta.push({
        source_file: item.source.name,
        period_start: dateToISO(item.period[0]),
        period_end: dateToISO(item.period[1]),
        sha256: await sha256(item.source.arrayBuffer),
        size_bytes: item.source.size,
        modified_at: fileModifiedISO(item.source),
      });
    }

    const sourceTotals = sourceMonthTotals(
      raw.basic_rows,
      raw.demand_rows,
      raw.dispatch_rows,
      analysisPeriod.months,
    );
    const cleaned = cleanTestUnits(raw.basic_rows, raw.demand_rows, raw.dispatch_rows);
    const atomic = buildAtomicTables(
      cleaned.basic_rows,
      cleaned.demand_rows,
      cleaned.dispatch_rows,
    );
    const supplierMonthly = aggregateSupplierMonthly(atomic.supplier_rows);
    const warehouseMonthly = aggregateWarehouseMonthly(atomic.unit_rows);
    const regions = orderedRegions(supplierMonthly.concat(warehouseMonthly));
    if (!regions.length) {
      throw new Error("清洗后没有可统计的区域数据，请检查源文件是否只有测试记录");
    }
    const geography = buildGeography(regions);
    const display = buildDisplayRows(supplierMonthly, warehouseMonthly, geography);
    const displayMetrics = buildDisplayMetrics(
      display.displayRows,
      supplierMonthly,
      analysisPeriod.months,
    );
    const warehouseDisplay = buildWarehouseDisplay(
      warehouseMonthly,
      supplierMonthly,
      display.regions,
      display.jobs,
      analysisPeriod.months,
      geography,
    );
    const reconciliation = monthRegionReconciliation(
      raw.basic_rows,
      raw.demand_rows,
      raw.dispatch_rows,
    );
    const quality = buildQuality(
      raw,
      cleaned,
      atomic,
      supplierMonthly,
      reconciliation,
    );
    quality.region_classification = {
      status: geography.unknown_regions.length ? "REVIEW" : "PASS",
      unknown_regions: geography.unknown_regions,
      mapped_region_count: geography.regions.length - geography.unknown_regions.length,
      total_region_count: geography.regions.length,
    };
    if (geography.unknown_regions.length) {
      quality.critical_checks.push({
        id: "region_classification",
        label: "区域大区归属",
        status: "REVIEW",
        difference_count: geography.unknown_regions.length,
      });
      quality.status = "REVIEW";
    }
    const coverage = coverageAudit(
      parsedFiles.map((item) => item.period),
      analysisPeriod,
    );
    const scope = scopeNames(geography, analysisPeriod);
    const supplierCount = new Set(
      supplierMonthly.map((row) => [row.region, row.supplier_id, row.supplier_name].join("\u0001")),
    ).size;
    const cleanTotals = {
      base_demand: cleanNumber(
        atomic.unit_rows.reduce((sum, row) => sum + row.raw_demand, 0),
      ),
      valid_supplement: cleanNumber(
        atomic.unit_rows.reduce((sum, row) => sum + row.supplement, 0),
      ),
      true_demand: cleanNumber(
        atomic.unit_rows.reduce((sum, row) => sum + row.true_demand, 0),
      ),
      issued: cleanNumber(atomic.unit_rows.reduce((sum, row) => sum + row.issued, 0)),
      dispatched: cleanNumber(
        atomic.unit_rows.reduce((sum, row) => sum + row.dispatched, 0),
      ),
      raw_effective: cleanNumber(
        atomic.supplier_rows.reduce((sum, row) => sum + row.raw_effective, 0),
      ),
      reasonable_responsibility: cleanNumber(
        atomic.supplier_rows.reduce(
          (sum, row) => sum + row.reasonable_responsibility,
          0,
        ),
      ),
      adjusted_effective: cleanNumber(
        atomic.supplier_rows.reduce((sum, row) => sum + row.adjusted_effective, 0),
      ),
      warehouse_effective: cleanNumber(
        atomic.unit_rows.reduce((sum, row) => sum + row.warehouse_effective, 0),
      ),
    };
    const scopeSummary = buildScopeSummary(
      displayMetrics,
      warehouseDisplay,
      atomic.unit_rows,
      atomic.supplier_rows,
      display.regions,
      analysisPeriod.months,
    );
    const regionSummary = buildRegionSummary(
      displayMetrics,
      warehouseDisplay,
      atomic.unit_rows,
      display.regions,
    );
    const marketSummary = buildMarketSummary(
      displayMetrics,
      warehouseDisplay,
      atomic.unit_rows,
      geography,
    );
    const comparisonSummary = marketSummary.concat(regionSummary);
    const periodSupplierSummary = displayMetrics
      .filter((row) => row.row_type === "supplier_total")
      .map((row) => ({
        market: row.market,
        region: row.region,
        supplier_id: row.supplier_id,
        supplier_name: row.supplier_name,
        raw_month_rates: row.by_month.map((month) => month.raw_rate),
        adjusted_month_rates: row.by_month.map((month) => month.adjusted_rate),
        raw_average: row.raw_average,
        adjusted_average: row.adjusted_average,
        raw_q2_weighted: row.raw_weighted,
        adjusted_q2_weighted: row.adjusted_weighted,
        raw_period_weighted: row.raw_weighted,
        adjusted_period_weighted: row.adjusted_weighted,
        average_lift:
          row.raw_average != null && row.adjusted_average != null
            ? row.adjusted_average - row.raw_average
            : null,
        total_issued: row.totals.issued,
        total_raw_effective: row.totals.raw_effective,
        total_responsibility: row.totals.reasonable_responsibility,
        total_adjusted_effective: row.totals.adjusted_effective,
        total_relief: row.totals.over_issue_relief,
      }))
      .sort((left, right) => {
        return (
          regionRank(left.region) - regionRank(right.region) ||
          (right.raw_average ?? -1) - (left.raw_average ?? -1) ||
          left.supplier_name.localeCompare(right.supplier_name, "zh-CN")
        );
      });

    return {
      meta: {
        title: scope.title,
        management_title: scope.management_title,
        report_name: scope.report_name,
        export_filename: scope.filename,
        generated_at: new Date().toISOString().slice(0, 19),
        scope: `${display.regions.join("、")}，${dateToISO(
          analysisPeriod.start,
        )}至${dateToISO(
          analysisPeriod.end,
        )}`,
        scope_prefix: scope.prefix,
        unit: "人日；比例分摊生成的合理责任为责任人日等值",
        source_file_count: raw.file_meta.length,
        markets: geography.markets,
        region_market: geography.region_market,
        unknown_regions: geography.unknown_regions,
        scope_total_label: geography.scope_total_label,
        regions: display.regions,
        jobs: display.jobs,
        supplier_count: supplierCount,
        months: analysisPeriod.months,
        month_count: analysisPeriod.months.length,
        quarter: analysisPeriod.shortCode,
        quarter_code: analysisPeriod.code,
        period_code: analysisPeriod.code,
        period_short: analysisPeriod.shortCode,
        period_label: analysisPeriod.label,
        period_start: dateToISO(analysisPeriod.start),
        period_end: dateToISO(analysisPeriod.end),
        primary_metric: "原始接单兑现率",
        comparison_metric: "超发校正履约率",
        no_data_rule:
          "当月分母为0时显示—且不参与统计期月均；有分母但无兑现显示0%",
        coverage,
      },
      method: {
        atomic_grain: "区域×仓库×组×日期×工种×班次",
        raw_rate: "Σmin(供应商派遣,供应商发单) ÷ Σ供应商发单",
        reasonable_responsibility:
          "供应商发单 × min(1,真实需求÷单元总发单)",
        adjusted_rate: "Σmin(供应商派遣,合理责任) ÷ Σ合理责任",
        warehouse_rate: "Σmin(单元总派遣,真实需求) ÷ Σ真实需求",
        issue_coverage: "Σmin(单元总发单,真实需求) ÷ Σ真实需求",
        average: "有有效分母月份的月度比率算术平均",
        period_weighted: "统计期累计有效兑现 ÷ 统计期累计分母",
        q2_weighted: "统计期累计有效兑现 ÷ 统计期累计分母",
        zero_demand_issue:
          "保留在原始接单兑现率；合理责任为0，不进入校正率和仓需分母",
        test_unit: "出现测试供应商的完整原子单元及同组日补员整体剔除",
      },
      source_files: raw.file_meta,
      source_totals: sourceTotals,
      clean_totals: cleanTotals,
      display_rows: display.displayRows,
      display_metrics: displayMetrics,
      supplier_monthly: supplierMonthly,
      supplier_period_summary: periodSupplierSummary,
      supplier_q2_summary: periodSupplierSummary,
      warehouse_monthly: warehouseMonthly,
      warehouse_display: warehouseDisplay,
      unit_atomic: atomic.unit_rows,
      supplier_atomic: atomic.supplier_rows,
      scope_summary: scopeSummary,
      region_summary: regionSummary,
      market_summary: marketSummary,
      comparison_summary: comparisonSummary,
      quality,
    };
  }

  return {
    analyzeFiles,
    parseDate,
    parseSpan,
    quarterFromDate,
    periodFromRange,
    cleanNumber,
    safeRate,
    monthAverage,
    marketForRegion,
    constants: {
      EPS,
      MARKET_ORDER,
      REGION_ORDER,
      REGION_MARKET,
      DEMAND_DIMS,
      SUPPLIER_DIMS,
    },
  };
});

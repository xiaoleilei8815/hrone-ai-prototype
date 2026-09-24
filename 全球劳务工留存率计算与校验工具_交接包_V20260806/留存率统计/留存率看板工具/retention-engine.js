(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.RetentionEngine = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const DAY_MS = 86400000;
  const FINAL_CONFIRM_STATUSES = new Set(["已确认", "已复核"]);
  const REGION_METADATA = {
    英国区: { globalRegion: "欧洲", country: "英国" },
    德国区: { globalRegion: "欧洲", country: "德国" },
    法国区: { globalRegion: "欧洲", country: "法国" },
    捷克区: { globalRegion: "欧洲", country: "捷克" },
    波兰区: { globalRegion: "欧洲", country: "波兰" },
    意大利区: { globalRegion: "欧洲", country: "意大利" },
    西班牙区: { globalRegion: "欧洲", country: "西班牙" },
    加州区: { globalRegion: "美洲", country: "美国" },
    新泽西区: { globalRegion: "美洲", country: "美国" },
    亚特兰大区: { globalRegion: "美洲", country: "美国" },
    芝加哥区: { globalRegion: "美洲", country: "美国" },
    休斯顿区: { globalRegion: "美洲", country: "美国" },
    达拉斯区: { globalRegion: "美洲", country: "美国" },
    萨凡纳区: { globalRegion: "美洲", country: "美国" },
    迈阿密区: { globalRegion: "美洲", country: "美国" },
    西雅图区: { globalRegion: "美洲", country: "美国" },
    诺福克区: { globalRegion: "美洲", country: "美国" },
    加拿大区: { globalRegion: "美洲", country: "加拿大" },
    澳洲区: { globalRegion: "亚太", country: "澳大利亚" },
    韩国区: { globalRegion: "亚太", country: "韩国" },
    日本区: { globalRegion: "亚太", country: "日本" },
  };
  const COUNTRY_MAP = Object.fromEntries(
    Object.entries(REGION_METADATA).map(([region, metadata]) => [region, metadata.country])
  );
  const GLOBAL_REGION_ORDER = ["欧洲", "美洲", "亚太", "未识别"];
  const SUPPLIER_ALIASES = new Map([
    ["niden service sp z o o", "niden"],
    ["niden", "niden"],
    ["dreman sp z o o", "dreman"],
    ["dreman", "dreman"],
    ["atlas work spo ka z ograniczona odpowiedzialnoscia", "atlaswork"],
    ["atlaswork", "atlaswork"],
    ["gi group sp z o o", "gi group"],
    ["gi group", "gi group"],
    ["ama service ug 安保", "ama service ug"],
    ["ama service ug", "ama service ug"],
    ["attal group s p a", "attal group s p a"],
    ["zenith zone sp z o o", "zenith"],
    ["zenith", "zenith"],
  ]);

  function cleanText(value) {
    return value == null ? "" : String(value).trim();
  }

  function normalizeSupplier(value) {
    let text = cleanText(value).normalize("NFKD").toLowerCase();
    text = text.replace(/[\u0300-\u036f]/g, "");
    text = text.replace(/[^0-9a-z\u4e00-\u9fff]+/g, " ").replace(/\s+/g, " ").trim();
    return SUPPLIER_ALIASES.get(text) || text;
  }

  function makeUtcDay(year, month, day) {
    return Math.floor(Date.UTC(year, month - 1, day) / DAY_MS);
  }

  function parseDateValue(value, XLSX) {
    if (value == null || value === "") return null;
    if (value instanceof Date && !Number.isNaN(value.valueOf())) {
      return makeUtcDay(value.getFullYear(), value.getMonth() + 1, value.getDate());
    }
    if (typeof value === "number" && Number.isFinite(value)) {
      const parsed = XLSX && XLSX.SSF ? XLSX.SSF.parse_date_code(value) : null;
      if (parsed && parsed.y && parsed.m && parsed.d) return makeUtcDay(parsed.y, parsed.m, parsed.d);
      return null;
    }
    const text = cleanText(value);
    const match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (match) return makeUtcDay(Number(match[1]), Number(match[2]), Number(match[3]));
    const parsed = new Date(text);
    if (!Number.isNaN(parsed.valueOf())) {
      return makeUtcDay(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate());
    }
    return null;
  }

  function isoToDay(value) {
    return parseDateValue(value, null);
  }

  function dayToISO(day) {
    if (day == null || !Number.isFinite(day)) return "";
    return new Date(day * DAY_MS).toISOString().slice(0, 10);
  }

  function monthsInRange(startDay, endDay) {
    if (startDay == null || endDay == null || startDay > endDay) return [];
    const start = new Date(startDay * DAY_MS);
    const end = new Date(endDay * DAY_MS);
    let year = start.getUTCFullYear();
    let month = start.getUTCMonth() + 1;
    const endYear = end.getUTCFullYear();
    const endMonth = end.getUTCMonth() + 1;
    const output = [];
    while (year < endYear || (year === endYear && month <= endMonth)) {
      output.push(`${year}-${String(month).padStart(2, "0")}`);
      month += 1;
      if (month > 12) {
        month = 1;
        year += 1;
      }
    }
    return output;
  }

  function lowerBound(values, target) {
    let low = 0;
    let high = values.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (values[middle] < target) low = middle + 1;
      else high = middle;
    }
    return low;
  }

  function upperBound(values, target) {
    let low = 0;
    let high = values.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (values[middle] <= target) low = middle + 1;
      else high = middle;
    }
    return low;
  }

  function valuesInRange(values, start, end) {
    if (!values || !values.length) return [];
    return values.slice(lowerBound(values, start), upperBound(values, end));
  }

  function lastAtOrBefore(values, end) {
    if (!values || !values.length) return null;
    const position = upperBound(values, end) - 1;
    return position >= 0 ? values[position] : null;
  }

  function findSheetWithHeaders(workbook, XLSX, requiredHeaders) {
    for (const sheetName of workbook.SheetNames || []) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet || !sheet["!ref"]) continue;
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      const maxHeaderRow = Math.min(range.e.r, range.s.r + 4);
      for (let row = range.s.r; row <= maxHeaderRow; row += 1) {
        const headers = [];
        const index = new Map();
        for (let column = range.s.c; column <= range.e.c; column += 1) {
          const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
          const header = cleanText(cell && cell.v);
          headers[column] = header;
          if (header) index.set(header, column);
        }
        if (requiredHeaders.every((header) => index.has(header))) {
          return { sheet, sheetName, headerRow: row, range, headers, index };
        }
      }
    }
    throw new Error(`找不到包含字段【${requiredHeaders.join("、")}】的工作表`);
  }

  function cellValue(sheet, XLSX, row, column) {
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
    return cell ? cell.v : null;
  }

  function relation(latestDispatch, latestEnd) {
    if (latestDispatch == null || latestEnd == null) return "一方为空";
    if (latestDispatch === latestEnd) return "L=E";
    return latestDispatch < latestEnd ? "L<E" : "L>E";
  }

  function parseManagementWorkbook(workbook, XLSX, fileName) {
    const required = ["工号", "首次派遣日期", "最晚派遣日期", "服务商名称", "区域"];
    const found = findSheetWithHeaders(workbook, XLSX, required);
    const { sheet, sheetName, headerRow, range, index } = found;
    const get = (row, header) => (index.has(header) ? cellValue(sheet, XLSX, row, index.get(header)) : null);
    const records = new Map();
    const duplicates = new Set();
    let minFirst = null;
    let maxFirst = null;
    let blankId = 0;
    for (let row = headerRow + 1; row <= range.e.r; row += 1) {
      const employeeId = cleanText(get(row, "工号"));
      if (!employeeId) {
        blankId += 1;
        continue;
      }
      if (records.has(employeeId)) duplicates.add(employeeId);
      const firstDispatch = parseDateValue(get(row, "首次派遣日期"), XLSX);
      const latestDispatch = parseDateValue(get(row, "最晚派遣日期"), XLSX);
      const latestEnd = parseDateValue(get(row, "最近一次派遣结束日期"), XLSX);
      if (firstDispatch != null) {
        minFirst = minFirst == null ? firstDispatch : Math.min(minFirst, firstDispatch);
        maxFirst = maxFirst == null ? firstDispatch : Math.max(maxFirst, firstDispatch);
      }
      const region = cleanText(get(row, "区域"));
      const regionMetadata = REGION_METADATA[region] || {
        globalRegion: "未识别",
        country: region || "未识别",
      };
      records.set(employeeId, {
        employeeId,
        name: cleanText(get(row, "姓名")),
        employeeStatus: cleanText(get(row, "状态")),
        dispatchStatus: cleanText(get(row, "派遣状态")),
        supplierName: cleanText(get(row, "服务商名称")),
        supplierCode: cleanText(get(row, "服务商编码")),
        region,
        globalRegion: regionMetadata.globalRegion,
        country: regionMetadata.country,
        firstDispatch,
        latestDispatch,
        latestEnd,
        endType: cleanText(get(row, "派遣结束类型")),
        endReason: cleanText(get(row, "派遣结束原因")),
        endDetail: cleanText(get(row, "派遣结束具体原因")),
      });
    }
    return {
      records,
      profile: {
        fileName,
        sheetName,
        rows: Math.max(0, range.e.r - headerRow),
        uniqueEmployees: records.size,
        duplicateEmployees: duplicates.size,
        blankId,
        firstDispatchMin: dayToISO(minFirst),
        firstDispatchMax: dayToISO(maxFirst),
      },
    };
  }

  function createAttendanceStore() {
    return {
      employees: new Map(),
      coverageMin: null,
      coverageMax: null,
      coverageDates: new Set(),
      coverageMonths: new Set(),
      profiles: [],
      finalized: false,
    };
  }

  function employeeEvidence(store, employeeId) {
    if (!store.employees.has(employeeId)) {
      store.employees.set(employeeId, {
        validDays: new Map(),
        recordDays: new Set(),
        unconfirmedDays: new Set(),
        validDayList: [],
        recordDayList: [],
        unconfirmedDayList: [],
      });
    }
    return store.employees.get(employeeId);
  }

  function mergeAttendanceWorkbook(workbook, XLSX, fileName, store, cohortIds) {
    const required = ["工号", "考勤日期", "时长总计", "确认状态", "考勤类型"];
    const found = findSheetWithHeaders(workbook, XLSX, required);
    const { sheet, sheetName, headerRow, range, index } = found;
    const get = (row, header) => (index.has(header) ? cellValue(sheet, XLSX, row, index.get(header)) : null);
    const uniqueEmployees = new Set();
    const cohortEmployees = new Set();
    const seenEmployeeDates = new Set();
    const duplicateEmployeeDates = new Set();
    let rows = 0;
    let cohortRows = 0;
    let validPositiveRows = 0;
    let positiveUnconfirmedRows = 0;
    let zeroHourRows = 0;
    let zeroHourFinalRows = 0;
    let dateMin = null;
    let dateMax = null;

    for (let row = headerRow + 1; row <= range.e.r; row += 1) {
      rows += 1;
      const employeeId = cleanText(get(row, "工号"));
      const attendanceDay = parseDateValue(get(row, "考勤日期"), XLSX);
      if (!employeeId || attendanceDay == null) continue;
      uniqueEmployees.add(employeeId);
      dateMin = dateMin == null ? attendanceDay : Math.min(dateMin, attendanceDay);
      dateMax = dateMax == null ? attendanceDay : Math.max(dateMax, attendanceDay);
      store.coverageMin = store.coverageMin == null ? attendanceDay : Math.min(store.coverageMin, attendanceDay);
      store.coverageMax = store.coverageMax == null ? attendanceDay : Math.max(store.coverageMax, attendanceDay);
      store.coverageDates.add(attendanceDay);
      store.coverageMonths.add(dayToISO(attendanceDay).slice(0, 7));

      const key = `${employeeId}|${attendanceDay}`;
      if (seenEmployeeDates.has(key)) duplicateEmployeeDates.add(key);
      else seenEmployeeDates.add(key);

      const totalHours = Number(get(row, "时长总计")) || 0;
      const confirmStatus = cleanText(get(row, "确认状态"));
      const valid = totalHours > 0 && FINAL_CONFIRM_STATUSES.has(confirmStatus);
      const unconfirmedPositive = totalHours > 0 && !FINAL_CONFIRM_STATUSES.has(confirmStatus);
      if (valid) validPositiveRows += 1;
      if (unconfirmedPositive) positiveUnconfirmedRows += 1;
      if (totalHours <= 0) {
        zeroHourRows += 1;
        if (FINAL_CONFIRM_STATUSES.has(confirmStatus)) zeroHourFinalRows += 1;
      }
      if (!cohortIds.has(employeeId)) continue;

      cohortRows += 1;
      cohortEmployees.add(employeeId);
      const evidence = employeeEvidence(store, employeeId);
      evidence.recordDays.add(attendanceDay);
      if (unconfirmedPositive) evidence.unconfirmedDays.add(attendanceDay);
      if (valid) {
        if (!evidence.validDays.has(attendanceDay)) {
          evidence.validDays.set(attendanceDay, {
            types: new Set(),
            suppliers: new Set(),
            supplierCodes: new Set(),
            regions: new Set(),
          });
        }
        const event = evidence.validDays.get(attendanceDay);
        event.types.add(cleanText(get(row, "考勤类型")));
        const supplier = cleanText(get(row, "供应商名称"));
        const supplierCode = cleanText(get(row, "供应商ID"));
        const region = cleanText(get(row, "区域"));
        if (supplier) event.suppliers.add(supplier);
        if (supplierCode) event.supplierCodes.add(supplierCode);
        if (region) event.regions.add(region);
      }
    }
    const profile = {
      fileName,
      sheetName,
      rows,
      cohortRows,
      dateMin: dayToISO(dateMin),
      dateMax: dayToISO(dateMax),
      uniqueEmployees: uniqueEmployees.size,
      cohortEmployees: cohortEmployees.size,
      validPositiveRows,
      positiveUnconfirmedRows,
      zeroHourRows,
      zeroHourFinalRows,
      duplicateEmployeeDates: duplicateEmployeeDates.size,
    };
    store.profiles.push(profile);
    store.finalized = false;
    return profile;
  }

  function finalizeAttendanceStore(store) {
    for (const evidence of store.employees.values()) {
      evidence.validDayList = Array.from(evidence.validDays.keys()).sort((a, b) => a - b);
      evidence.recordDayList = Array.from(evidence.recordDays).sort((a, b) => a - b);
      evidence.unconfirmedDayList = Array.from(evidence.unconfirmedDays).sort((a, b) => a - b);
    }
    store.finalized = true;
    return store;
  }

  function wilson(successes, total, z) {
    if (!total) return [null, null];
    const score = z || 1.96;
    const p = successes / total;
    const denominator = 1 + (score * score) / total;
    const center = (p + (score * score) / (2 * total)) / denominator;
    const margin =
      (score * Math.sqrt((p * (1 - p)) / total + (score * score) / (4 * total * total))) / denominator;
    return [Math.max(0, center - margin), Math.min(1, center + margin)];
  }

  function calculate(cohortRecords, attendanceStore, latestRecords, options) {
    if (!attendanceStore.finalized) finalizeAttendanceStore(attendanceStore);
    const periodDays = Math.max(1, Number(options.periodDays) || 90);
    const windowDays = Math.max(1, Number(options.windowDays) || 14);
    const cohortStart = options.cohortStart == null ? null : Number(options.cohortStart);
    const cohortEnd = options.cohortEnd == null ? null : Number(options.cohortEnd);
    const overrides = options.overrides || new Map();
    const results = [];

    for (const record of cohortRecords.values()) {
      const firstDispatch = record.firstDispatch;
      if (firstDispatch == null) continue;
      if (cohortStart != null && firstDispatch < cohortStart) continue;
      if (cohortEnd != null && firstDispatch > cohortEnd) continue;

      const observationDay = firstDispatch + periodDays;
      const windowStart = observationDay - (windowDays - 1);
      const evidence = attendanceStore.employees.get(record.employeeId) || {
        validDayList: [],
        recordDayList: [],
        unconfirmedDayList: [],
        validDays: new Map(),
      };
      const validWindowDays = valuesInRange(evidence.validDayList, windowStart, observationDay);
      const recordWindowDays = valuesInRange(evidence.recordDayList, windowStart, observationDay);
      const unconfirmedWindowDays = valuesInRange(evidence.unconfirmedDayList, windowStart, observationDay);
      const lastValidBeforeT = lastAtOrBefore(evidence.validDayList, observationDay);
      const firstValid = evidence.validDayList.length ? evidence.validDayList[0] : null;
      const windowTypes = new Set();
      const windowSuppliers = new Set();
      const windowSupplierCodes = new Set();
      const windowRegions = new Set();
      for (const day of validWindowDays) {
        const event = evidence.validDays.get(day);
        if (!event) continue;
        for (const value of event.types) windowTypes.add(value);
        for (const value of event.suppliers) windowSuppliers.add(value);
        for (const value of event.supplierCodes) windowSupplierCodes.add(value);
        for (const value of event.regions) windowRegions.add(value);
      }

      const requiredCoverageMonths = monthsInRange(windowStart, observationDay);
      const missingCoverageMonths = requiredCoverageMonths.filter(
        (month) => !attendanceStore.coverageMonths.has(month)
      );
      const dataMature =
        attendanceStore.coverageMin != null &&
        attendanceStore.coverageMax != null &&
        windowStart >= attendanceStore.coverageMin &&
        observationDay <= attendanceStore.coverageMax &&
        missingCoverageMonths.length === 0;
      let missingGlobalCoverageDays = 0;
      if (dataMature) {
        for (let day = windowStart; day <= observationDay; day += 1) {
          if (!attendanceStore.coverageDates.has(day)) missingGlobalCoverageDays += 1;
        }
      }

      const latestDispatch = record.latestDispatch;
      const latestEnd = record.latestEnd;
      const endedBeforeTWithoutReturn =
        latestEnd != null &&
        latestEnd < observationDay &&
        (lastValidBeforeT == null || lastValidBeforeT <= latestEnd);
      let modelResult;
      let reason;
      if (!dataMature) {
        modelResult = "未成熟";
        reason = missingCoverageMonths.length
          ? `考勤源缺少月份：${missingCoverageMonths.join("、")}`
          : "观察期或考勤窗口尚未被数据完整覆盖";
      } else if (latestDispatch == null) {
        modelResult = "待核验";
        reason = "缺少最晚派遣日期";
      } else if (latestDispatch < observationDay) {
        modelResult = "未留存";
        reason = "最晚派遣日期早于T";
      } else if (endedBeforeTWithoutReturn) {
        modelResult = "未留存";
        reason = "T前结束且无结束后有效出勤";
      } else if (!validWindowDays.length) {
        modelResult = "未留存";
        reason = `T前${windowDays}日无有效出勤`;
      } else {
        modelResult = "留存";
        reason = `通过${periodDays}天留存`;
      }

      const latestCrossRecord = latestRecords ? latestRecords.get(record.employeeId) : null;
      const latestMatch = Boolean(latestCrossRecord);
      const crossLatestDispatch = latestCrossRecord ? latestCrossRecord.latestDispatch : null;
      const crossLatestEnd = latestCrossRecord ? latestCrossRecord.latestEnd : null;
      const latestFieldMismatch =
        latestMatch && (crossLatestDispatch !== latestDispatch || crossLatestEnd !== latestEnd);
      const blindEliminationRisk = latestMatch && latestDispatch != null && latestDispatch >= observationDay;
      const leRelation = relation(latestDispatch, latestEnd);
      const normalizedWindowSuppliers = new Set(Array.from(windowSuppliers, normalizeSupplier));
      const supplierMismatch =
        normalizedWindowSuppliers.size > 0 && !normalizedWindowSuppliers.has(normalizeSupplier(record.supplierName));
      const manualFlags = [];
      const qualityFlags = [];
      if (latestDispatch === observationDay || latestEnd === observationDay) {
        manualFlags.push("T边界日期需确认");
        qualityFlags.push("日期恰好等于T");
      }
      if (leRelation === "L<E") {
        qualityFlags.push("L<E（存在派遣删除异常）");
        if (latestDispatch < observationDay && observationDay <= latestEnd) manualFlags.push("L<E且T落在冲突区间");
      }
      if (latestDispatch != null && latestDispatch < observationDay && lastValidBeforeT > latestDispatch) {
        manualFlags.push("L日之后仍有有效出勤");
        qualityFlags.push("系统L与实际考勤冲突");
      }
      if (latestFieldMismatch) {
        manualFlags.push("交叉表L/E与主表不一致");
        qualityFlags.push("两份用工管理表L/E不一致");
      }
      if (dataMature && latestDispatch >= observationDay && !validWindowDays.length) {
        manualFlags.push("系统派遣覆盖T但窗口无有效出勤");
        qualityFlags.push("系统派遣覆盖T但窗口无有效出勤");
      }
      if (unconfirmedWindowDays.length) {
        manualFlags.push("窗口存在未确认正工时");
        qualityFlags.push("窗口存在未确认正工时");
      }
      if (!recordWindowDays.length && dataMature) qualityFlags.push("T窗口无考勤记录");
      if (!evidence.validDayList.length) qualityFlags.push("全部考勤源无有效出勤");
      if (firstValid != null && firstValid - firstDispatch > 7) qualityFlags.push("首次有效出勤晚于首次派遣7天以上");
      if (supplierMismatch) {
        manualFlags.push("T窗口供应商与入职供应商不一致");
        qualityFlags.push("T窗口供应商与入职供应商不一致");
      }
      if (normalizedWindowSuppliers.size > 1) manualFlags.push("T窗口出现多个供应商");
      if (missingGlobalCoverageDays > 0) qualityFlags.push(`全局考勤日期缺口${missingGlobalCoverageDays}天`);
      if (missingCoverageMonths.length) qualityFlags.push(`考勤源缺少月份：${missingCoverageMonths.join("、")}`);

      const uniqueManualFlags = Array.from(new Set(manualFlags));
      const override = overrides.get(record.employeeId) || null;
      const finalResult = override && override.result ? override.result : modelResult;
      results.push({
        ...record,
        periodDays,
        windowDays,
        observationDay,
        windowStart,
        latestMatch,
        crossLatestDispatch,
        crossLatestEnd,
        latestFieldMismatch,
        blindEliminationRisk,
        latestEndRelation: leRelation,
        firstValid,
        firstValidGapDays: firstValid == null ? null : firstValid - firstDispatch,
        lastValidBeforeT,
        validWindowDays,
        recordWindowDays: recordWindowDays.length,
        unconfirmedWindowDays,
        windowTypes: Array.from(windowTypes).sort(),
        windowSuppliers: Array.from(windowSuppliers).sort(),
        windowSupplierCodes: Array.from(windowSupplierCodes).sort(),
        windowRegions: Array.from(windowRegions).sort(),
        supplierMismatch,
        dataMature,
        missingGlobalCoverageDays,
        missingCoverageMonths,
        modelResult,
        reason,
        manualReview: uniqueManualFlags.length > 0,
        manualFlags: uniqueManualFlags,
        qualityFlags: Array.from(new Set(qualityFlags)),
        overrideResult: override ? override.result || "" : "",
        overrideNote: override ? override.note || "" : "",
        finalResult,
        entryMonth: dayToISO(firstDispatch).slice(0, 7),
      });
    }
    return results;
  }

  function aggregateGroup(rows, keys) {
    const groups = new Map();
    for (const row of rows) {
      const key = keys.map((name) => row[name] || "").join("\u0001");
      if (!groups.has(key)) {
        const identity = {};
        for (const name of keys) identity[name] = row[name] || "";
        groups.set(key, { ...identity, rows: [] });
      }
      groups.get(key).rows.push(row);
    }
    const output = [];
    for (const group of groups.values()) {
      const cohort = group.rows.length;
      const immature = group.rows.filter((row) => row.finalResult === "未成熟").length;
      const pending = group.rows.filter((row) => row.finalResult === "待核验").length;
      const exempt = group.rows.filter((row) => row.finalResult === "豁免").length;
      const retained = group.rows.filter((row) => row.finalResult === "留存").length;
      const notRetained = group.rows.filter((row) => row.finalResult === "未留存").length;
      const denominator = retained + notRetained;
      const [ciLow, ciHigh] = wilson(retained, denominator);
      const summary = {
        ...Object.fromEntries(Object.entries(group).filter(([key]) => key !== "rows")),
        cohort,
        immature,
        pending,
        exempt,
        denominator,
        retained,
        notRetained,
        retentionRate: denominator ? retained / denominator : null,
        ciLow,
        ciHigh,
        manualReview: group.rows.filter((row) => row.manualReview).length,
        blindMatches: group.rows.filter((row) => row.latestMatch).length,
        blindEliminationRisk: group.rows.filter((row) => row.blindEliminationRisk).length,
        latestBeforeT: group.rows.filter((row) => row.reason === "最晚派遣日期早于T").length,
        endedBeforeT: group.rows.filter((row) => row.reason === "T前结束且无结束后有效出勤").length,
        noValidAttendance: group.rows.filter((row) => /^T前\d+日无有效出勤$/.test(row.reason)).length,
        sampleFlag: denominator < 10 ? "样本不足（<10）" : "可比较",
        months: {},
      };
      for (const month of new Set(group.rows.map((row) => row.entryMonth))) {
        const monthRows = group.rows.filter((row) => row.entryMonth === month);
        const monthRetained = monthRows.filter((row) => row.finalResult === "留存").length;
        const monthNotRetained = monthRows.filter((row) => row.finalResult === "未留存").length;
        const monthDenominator = monthRetained + monthNotRetained;
        summary.months[month] = {
          cohort: monthRows.length,
          denominator: monthDenominator,
          retained: monthRetained,
          rate: monthDenominator ? monthRetained / monthDenominator : null,
        };
      }
      output.push(summary);
    }
    return output;
  }

  function aggregate(results) {
    const retained = results.filter((row) => row.finalResult === "留存").length;
    const notRetained = results.filter((row) => row.finalResult === "未留存").length;
    const denominator = retained + notRetained;
    const globalRegionRank = (value) => {
      const index = GLOBAL_REGION_ORDER.indexOf(value);
      return index >= 0 ? index : GLOBAL_REGION_ORDER.length;
    };
    const globalRegions = aggregateGroup(results, ["globalRegion"]).sort((a, b) => {
      return globalRegionRank(a.globalRegion) - globalRegionRank(b.globalRegion);
    });
    const countryTotals = aggregateGroup(results, ["globalRegion", "country"]).sort((a, b) => {
      const regionRank = globalRegionRank(a.globalRegion) - globalRegionRank(b.globalRegion);
      if (regionRank) return regionRank;
      return a.country.localeCompare(b.country, "zh-CN");
    });
    const operatingRegions = aggregateGroup(results, ["globalRegion", "country", "region"]).sort((a, b) => {
      if (b.retentionRate !== a.retentionRate) return (b.retentionRate ?? -1) - (a.retentionRate ?? -1);
      return b.denominator - a.denominator;
    });
    const suppliers = aggregateGroup(results, ["globalRegion", "country", "region", "supplierName", "supplierCode"]).sort((a, b) => {
      const regionRank = globalRegionRank(a.globalRegion) - globalRegionRank(b.globalRegion);
      if (regionRank) return regionRank;
      if (a.country !== b.country) return a.country.localeCompare(b.country, "zh-CN");
      if (a.region !== b.region) return a.region.localeCompare(b.region, "zh-CN");
      if (b.retentionRate !== a.retentionRate) return (b.retentionRate ?? -1) - (a.retentionRate ?? -1);
      return b.denominator - a.denominator;
    });
    const reasons = new Map();
    for (const row of results) reasons.set(row.reason, (reasons.get(row.reason) || 0) + 1);
    return {
      cohort: results.length,
      immature: results.filter((row) => row.finalResult === "未成熟").length,
      pending: results.filter((row) => row.finalResult === "待核验").length,
      exempt: results.filter((row) => row.finalResult === "豁免").length,
      denominator,
      retained,
      notRetained,
      retentionRate: denominator ? retained / denominator : null,
      manualReview: results.filter((row) => row.manualReview).length,
      blindEliminationRisk: results.filter((row) => row.blindEliminationRisk).length,
      globalRegions,
      countryTotals,
      operatingRegions,
      countries: operatingRegions,
      suppliers,
      reasons: Array.from(reasons, ([reason, count]) => ({ reason, count })).sort((a, b) => b.count - a.count),
    };
  }

  return {
    DAY_MS,
    FINAL_CONFIRM_STATUSES,
    REGION_METADATA,
    COUNTRY_MAP,
    cleanText,
    normalizeSupplier,
    parseDateValue,
    isoToDay,
    dayToISO,
    monthsInRange,
    lowerBound,
    upperBound,
    valuesInRange,
    lastAtOrBefore,
    relation,
    parseManagementWorkbook,
    createAttendanceStore,
    mergeAttendanceWorkbook,
    finalizeAttendanceStore,
    calculate,
    aggregate,
    wilson,
  };
});

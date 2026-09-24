(function () {
  "use strict";

  const Engine = window.RetentionEngine;
  const XLSX = window.XLSX;
  const ReportExporter = window.RetentionReportExporter;
  const $ = (id) => document.getElementById(id);
  const nf = new Intl.NumberFormat("zh-CN");
  const GLOBAL_REGION_ORDER = ["欧洲", "美洲", "亚太", "未识别"];
  const state = {
    cohortMap: null,
    latestMap: new Map(),
    attendanceStore: null,
    profiles: { cohort: null, latest: null, attendance: [] },
    allResults: [],
    results: [],
    aggregate: null,
    scope: "",
    scopeInitialized: false,
    overrides: new Map(),
    periodDays: 90,
    windowDays: 14,
    peoplePage: 1,
    pageSize: 50,
    supplierSort: { key: "retentionRate", direction: "desc" },
    loading: false,
    downloadUrl: null,
  };

  const dom = {
    cohortFile: $("cohortFile"),
    latestFile: $("latestFile"),
    attendanceFiles: $("attendanceFiles"),
    cohortFileName: $("cohortFileName"),
    latestFileName: $("latestFileName"),
    attendanceFileName: $("attendanceFileName"),
    loadBtn: $("loadBtn"),
    calculateBtn: $("calculateBtn"),
    exportBtn: $("exportBtn"),
    downloadReadyLink: $("downloadReadyLink"),
    resetBtn: $("resetBtn"),
    progressBar: $("progressBar"),
    progressText: $("progressText"),
    systemMessage: $("systemMessage"),
    cohortStart: $("cohortStart"),
    cohortEnd: $("cohortEnd"),
    scopeSelect: $("scopeSelect"),
    periodCustom: $("periodCustom"),
    periodCustomWrap: $("periodCustomWrap"),
    windowCustom: $("windowCustom"),
    windowCustomWrap: $("windowCustomWrap"),
    coverageHint: $("coverageHint"),
  };

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function uiText(value) {
    return String(value == null ? "" : value)
      .replace(/L<T/g, "最晚派遣早于截止日")
      .replace(/E<T/g, "最近结束早于截止日")
      .replace(/T前/g, "截止日前")
      .replace(/T窗口/g, "截止日校验期")
      .replace(/覆盖T/g, "覆盖截止日")
      .replace(/等于T/g, "等于截止日")
      .replace(/T落在/g, "截止日落在")
      .replace(/T边界/g, "截止日边界")
      .replace(/T日/g, "截止日")
      .replace(/个人T/g, "个人截止日")
      .replace(/\bT\b/g, "截止日");
  }

  function orderedGlobalRegions(rows) {
    return Array.from(new Set(rows.map((row) => row.globalRegion || "未识别"))).sort((left, right) => {
      const leftIndex = GLOBAL_REGION_ORDER.indexOf(left);
      const rightIndex = GLOBAL_REGION_ORDER.indexOf(right);
      const leftRank = leftIndex >= 0 ? leftIndex : GLOBAL_REGION_ORDER.length;
      const rightRank = rightIndex >= 0 ? rightIndex : GLOBAL_REGION_ORDER.length;
      return leftRank - rightRank || left.localeCompare(right, "zh-CN");
    });
  }

  function combinedScopeLabel(regions) {
    const recognized = regions.filter((region) => region !== "未识别");
    if (recognized.length === 1 && regions.length === 1) return recognized[0];
    if (recognized.length === 2 && recognized.includes("美洲") && recognized.includes("亚太") && regions.length === 2) {
      return "美洲亚太";
    }
    return "全球";
  }

  function currentScopeLabel() {
    return state.scope || combinedScopeLabel(orderedGlobalRegions(state.allResults));
  }

  function day(value) {
    return value == null ? "" : Engine.dayToISO(value);
  }

  function percent(value, digits) {
    if (value == null || !Number.isFinite(value)) return "—";
    return `${(value * 100).toFixed(digits == null ? 1 : digits)}%`;
  }

  function ciText(row) {
    if (row.ciLow == null || row.ciHigh == null) return "—";
    return `${percent(row.ciLow)} ~ ${percent(row.ciHigh)}`;
  }

  function statusClass(value) {
    return {
      留存: "retained",
      未留存: "not-retained",
      未成熟: "immature",
      待核验: "pending",
      豁免: "exempt",
    }[value] || "pending";
  }

  function statusBadge(value) {
    return `<span class="status ${statusClass(value)}">${escapeHtml(value || "待核验")}</span>`;
  }

  function setMessage(message, type) {
    dom.systemMessage.textContent = message;
    dom.systemMessage.className = `system-message ${type || "neutral"}`;
  }

  function setProgress(value, text) {
    const bounded = Math.max(0, Math.min(100, Number(value) || 0));
    dom.progressBar.style.width = `${bounded}%`;
    dom.progressText.textContent = text;
  }

  function setBusy(busy) {
    state.loading = busy;
    dom.loadBtn.disabled = busy;
    dom.loadBtn.textContent = busy ? "正在读取..." : "读取并校验数据";
    dom.calculateBtn.disabled = busy || !state.cohortMap || !state.attendanceStore;
    dom.exportBtn.disabled = busy || !state.aggregate;
    dom.scopeSelect.disabled = busy || !state.allResults.length;
  }

  function clearDownloadArtifact() {
    if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
    state.downloadUrl = null;
    dom.downloadReadyLink.classList.add("hidden");
    dom.downloadReadyLink.removeAttribute("href");
    dom.downloadReadyLink.removeAttribute("download");
  }

  function exposeDownloadArtifact(blob, fileName) {
    clearDownloadArtifact();
    const url = ReportExporter.download(blob, fileName);
    state.downloadUrl = url;
    dom.downloadReadyLink.href = url;
    dom.downloadReadyLink.download = fileName;
    dom.downloadReadyLink.classList.remove("hidden");
  }

  function nextPaint() {
    return new Promise((resolve) => window.requestAnimationFrame(() => window.setTimeout(resolve, 0)));
  }

  async function readWorkbook(file) {
    const buffer = await file.arrayBuffer();
    return XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      cellFormula: false,
      cellHTML: false,
      cellNF: false,
      cellStyles: false,
    });
  }

  function updateFileLabel(input, label, multiple) {
    const files = Array.from(input.files || []);
    label.textContent = files.length
      ? multiple
        ? `${files.length}个文件：${files.map((file) => file.name).join("、")}`
        : files[0].name
      : "尚未选择";
    const control = input.closest(".upload-control");
    if (control) control.classList.toggle("loaded", files.length > 0);
  }

  function selectedDays(containerId, customInput) {
    const active = document.querySelector(`#${containerId} button.active`);
    if (!active) return null;
    if (active.dataset.days === "custom") return Number(customInput.value);
    return Number(active.dataset.days);
  }

  function syncParameters() {
    state.periodDays = Math.max(1, selectedDays("periodSegments", dom.periodCustom) || 90);
    state.windowDays = Math.max(1, selectedDays("windowSegments", dom.windowCustom) || 14);
    updateCoverageHint();
  }

  function configureSegments(containerId, customWrap, customInput) {
    const container = $(containerId);
    container.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-days]");
      if (!button) return;
      for (const sibling of container.querySelectorAll("button")) sibling.classList.remove("active");
      button.classList.add("active");
      customWrap.classList.toggle("hidden", button.dataset.days !== "custom");
      if (button.dataset.days === "custom") customInput.focus();
      syncParameters();
    });
    customInput.addEventListener("input", () => {
      const value = Number(customInput.value);
      if (Number.isFinite(value) && value >= 1) syncParameters();
    });
  }

  function dateRangeFromCohort() {
    if (!state.cohortMap) return { min: null, max: null, count: 0 };
    const start = Engine.isoToDay(dom.cohortStart.value);
    const end = Engine.isoToDay(dom.cohortEnd.value);
    let min = null;
    let max = null;
    let count = 0;
    for (const record of state.cohortMap.values()) {
      if (record.firstDispatch == null) continue;
      if (start != null && record.firstDispatch < start) continue;
      if (end != null && record.firstDispatch > end) continue;
      min = min == null ? record.firstDispatch : Math.min(min, record.firstDispatch);
      max = max == null ? record.firstDispatch : Math.max(max, record.firstDispatch);
      count += 1;
    }
    return { min, max, count };
  }

  function updateCoverageHint() {
    const range = dateRangeFromCohort();
    if (!range.count) {
      dom.coverageHint.textContent = state.cohortMap
        ? "当前日期范围内没有队列人员。"
        : "上传队列后将自动计算所需考勤覆盖范围。";
      return;
    }
    const requiredStart = range.min + state.periodDays - (state.windowDays - 1);
    const requiredEnd = range.max + state.periodDays;
    let message = `${range.count}人｜需要考勤覆盖 ${day(requiredStart)} 至 ${day(requiredEnd)}`;
    if (state.attendanceStore && state.attendanceStore.coverageMin != null) {
      const actualStart = state.attendanceStore.coverageMin;
      const actualEnd = state.attendanceStore.coverageMax;
      const sufficient = actualStart <= requiredStart && actualEnd >= requiredEnd;
      const requiredMonths = Engine.monthsInRange(requiredStart, requiredEnd);
      const missingMonths = requiredMonths.filter((month) => !state.attendanceStore.coverageMonths.has(month));
      message += `｜已上传 ${day(actualStart)} 至 ${day(actualEnd)}（${sufficient && !missingMonths.length ? "整体覆盖" : "存在未成熟队列"}）`;
      if (missingMonths.length) message += `｜缺少月份 ${missingMonths.join("、")}`;
    }
    dom.coverageHint.textContent = message;
  }

  function validateDateInputs() {
    const start = Engine.isoToDay(dom.cohortStart.value);
    const end = Engine.isoToDay(dom.cohortEnd.value);
    if (start == null || end == null) throw new Error("请设置完整的队列开始和结束日期");
    if (start > end) throw new Error("队列开始日期不能晚于结束日期");
    if (state.periodDays < 1 || state.windowDays < 1) throw new Error("观察周期和考勤窗口都必须至少为1天");
    return { start, end };
  }

  function updateScopeOptions() {
    const regions = orderedGlobalRegions(state.allResults);
    if (!state.scopeInitialized) {
      state.scope = regions.length === 1 ? regions[0] : "";
      state.scopeInitialized = true;
    } else if (state.scope && !regions.includes(state.scope)) {
      state.scope = "";
    }
    const allLabel = combinedScopeLabel(regions);
    dom.scopeSelect.innerHTML = `<option value="">全部已上传大区（${escapeHtml(allLabel)}）</option>${regions
      .map((region) => `<option value="${escapeHtml(region)}">${escapeHtml(region)}</option>`)
      .join("")}`;
    dom.scopeSelect.value = state.scope;
    dom.scopeSelect.disabled = !regions.length;
  }

  function applyScopeResults() {
    state.results = state.scope
      ? state.allResults.filter((row) => row.globalRegion === state.scope)
      : state.allResults.slice();
    state.aggregate = Engine.aggregate(state.results);
    state.peoplePage = 1;
    renderAll();
    dom.exportBtn.disabled = !state.results.length;
  }

  async function loadData() {
    if (!XLSX || !Engine) {
      setMessage("Excel解析模块未成功加载，请确认vendor目录完整。", "error");
      return;
    }
    const cohortFile = dom.cohortFile.files && dom.cohortFile.files[0];
    const latestFile = dom.latestFile.files && dom.latestFile.files[0];
    const attendanceFiles = Array.from(dom.attendanceFiles.files || []);
    if (!cohortFile || !attendanceFiles.length) {
      setMessage("请至少上传首次派遣队列和覆盖观察窗口的考勤文件。", "warning");
      return;
    }

    setBusy(true);
    setProgress(2, "开始读取首次派遣队列");
    setMessage("文件全部在当前浏览器内解析；大型考勤文件可能需要数十秒。", "neutral");
    try {
      state.overrides.clear();
      state.allResults = [];
      state.results = [];
      state.aggregate = null;
      state.scope = "";
      state.scopeInitialized = false;

      let workbook = await readWorkbook(cohortFile);
      const cohort = Engine.parseManagementWorkbook(workbook, XLSX, cohortFile.name);
      state.cohortMap = cohort.records;
      state.profiles.cohort = cohort.profile;
      workbook = null;

      if (!cohort.records.size) throw new Error("首次派遣队列中未读取到有效工号");
      dom.cohortStart.value = cohort.profile.firstDispatchMin;
      dom.cohortEnd.value = cohort.profile.firstDispatchMax;
      setProgress(12, `队列读取完成：${nf.format(cohort.records.size)}个唯一工号`);
      await nextPaint();

      state.latestMap = new Map();
      state.profiles.latest = null;
      if (latestFile) {
        workbook = await readWorkbook(latestFile);
        const latest = Engine.parseManagementWorkbook(workbook, XLSX, latestFile.name);
        state.latestMap = latest.records;
        state.profiles.latest = latest.profile;
        workbook = null;
        setProgress(18, `交叉表读取完成：${nf.format(latest.records.size)}个唯一工号`);
        await nextPaint();
      }

      const store = Engine.createAttendanceStore();
      const cohortIds = new Set(state.cohortMap.keys());
      state.profiles.attendance = [];
      for (let index = 0; index < attendanceFiles.length; index += 1) {
        const file = attendanceFiles[index];
        const base = latestFile ? 18 : 12;
        const span = 76;
        setProgress(base + (span * index) / attendanceFiles.length, `正在解析考勤 ${index + 1}/${attendanceFiles.length}：${file.name}`);
        await nextPaint();
        workbook = await readWorkbook(file);
        const profile = Engine.mergeAttendanceWorkbook(workbook, XLSX, file.name, store, cohortIds);
        state.profiles.attendance.push(profile);
        workbook = null;
        setProgress(base + (span * (index + 1)) / attendanceFiles.length, `已读取考勤 ${index + 1}/${attendanceFiles.length}`);
        await nextPaint();
      }
      state.attendanceStore = Engine.finalizeAttendanceStore(store);
      syncParameters();
      calculate();
      setProgress(100, `数据就绪：${nf.format(state.allResults.length)}人已完成逐人计算`);

      const duplicates = cohort.profile.duplicateEmployees;
      const coverage = `${day(store.coverageMin)} 至 ${day(store.coverageMax)}`;
      const regions = orderedGlobalRegions(state.allResults);
      const unknownCount = state.allResults.filter((row) => row.globalRegion === "未识别").length;
      setMessage(
        `读取完成。队列${nf.format(cohort.records.size)}人，识别大区：${regions.join("、") || "无"}；考勤覆盖${coverage}${duplicates ? `；队列存在${duplicates}个重复工号，已保留最后一条并列入质量审查` : ""}${unknownCount ? `；另有${nf.format(unknownCount)}人无法映射大区，请先核对区域字段` : ""}。`,
        duplicates || unknownCount ? "warning" : "success"
      );
    } catch (error) {
      console.error(error);
      setProgress(0, "读取失败");
      setMessage(`读取失败：${error.message || error}`, "error");
      state.aggregate = null;
    } finally {
      setBusy(false);
    }
  }

  function calculate() {
    if (!state.cohortMap || !state.attendanceStore) return;
    clearDownloadArtifact();
    syncParameters();
    try {
      const range = validateDateInputs();
      state.allResults = Engine.calculate(state.cohortMap, state.attendanceStore, state.latestMap, {
        periodDays: state.periodDays,
        windowDays: state.windowDays,
        cohortStart: range.start,
        cohortEnd: range.end,
        overrides: state.overrides,
      });
      updateScopeOptions();
      applyScopeResults();
      dom.calculateBtn.disabled = false;
      dom.scopeSelect.disabled = false;
    } catch (error) {
      setMessage(`计算失败：${error.message || error}`, "error");
    }
  }

  function renderAll() {
    renderDashboard();
    populateCountryFilters();
    renderSuppliers();
    renderPeople();
    renderReview();
    updateCoverageHint();
    $("countryChartTitle").textContent = `${currentScopeLabel()}成熟队列${state.periodDays}天留存率`;
  }

  function renderDashboard() {
    const summary = state.aggregate;
    if (!summary) return;
    $("kpiCohort").textContent = nf.format(summary.cohort);
    $("kpiDenominator").textContent = nf.format(summary.denominator);
    $("kpiRetained").textContent = nf.format(summary.retained);
    $("kpiNotRetained").textContent = nf.format(summary.notRetained);
    $("kpiRate").textContent = percent(summary.retentionRate);
    $("kpiImmature").textContent = nf.format(summary.immature);
    $("kpiReview").textContent = nf.format(summary.manualReview);

    const maturityNotice = $("maturityNotice");
    maturityNotice.classList.toggle("hidden", summary.immature === 0 && summary.pending === 0);
    maturityNotice.textContent = summary.immature
      ? `${nf.format(summary.immature)}人的观察期截止日或完整考勤校验期超出已上传数据覆盖，已按右删失处理并排除出计分分母。`
      : `${nf.format(summary.pending)}人因缺失最晚派遣日期暂不计分。`;

    const countries = summary.operatingRegions;
    $("countryCount").textContent = `${countries.length}个运营区域`;
    const countryBars = $("countryBars");
    if (!countries.length) {
      countryBars.className = "bar-chart empty-state";
      countryBars.textContent = "暂无结果";
    } else {
      countryBars.className = "bar-chart";
      countryBars.innerHTML = countries
        .map(
          (row) => `<div class="bar-row"><span class="bar-label" title="${escapeHtml(`${row.globalRegion}｜${row.country}｜${row.region}`)}">${escapeHtml(row.region)}</span><span class="bar-track"><span style="width:${Math.max(0, (row.retentionRate || 0) * 100)}%"></span></span><strong class="bar-value">${percent(row.retentionRate)}</strong></div>`
        )
        .join("");
    }

    const reasonBars = $("reasonBars");
    const reasonCounts = new Map();
    for (const row of state.results.filter((item) => item.finalResult === "未留存")) {
      const reason = row.modelResult === "未留存" ? row.reason : "人工改判为未留存";
      reasonCounts.set(reason, (reasonCounts.get(reason) || 0) + 1);
    }
    const reasons = Array.from(reasonCounts, ([reason, count]) => ({ reason, count })).sort((a, b) => b.count - a.count);
    if (!reasons.length) {
      reasonBars.className = "reason-chart empty-state";
      reasonBars.textContent = "暂无未留存原因";
    } else {
      const max = Math.max(...reasons.map((item) => item.count));
      reasonBars.className = "reason-chart";
      reasonBars.innerHTML = reasons
        .slice(0, 7)
        .map(
          (item) => `<div class="bar-row reason"><span class="bar-label" title="${escapeHtml(uiText(item.reason))}">${escapeHtml(uiText(item.reason))}</span><span class="bar-track"><span style="width:${(item.count / max) * 100}%"></span></span><strong class="bar-value">${nf.format(item.count)}</strong></div>`
        )
        .join("");
    }

    const areaKey = (row) => `${row.globalRegion}\u0001${row.country}\u0001${row.region}`;
    const supplierCount = new Map();
    for (const row of summary.suppliers) supplierCount.set(areaKey(row), (supplierCount.get(areaKey(row)) || 0) + 1);
    $("countryTableBody").innerHTML = countries.length
      ? countries
          .map(
            (row) => `<tr><td>${escapeHtml(row.globalRegion)}</td><td>${escapeHtml(row.country)}</td><td>${escapeHtml(row.region)}</td><td class="number">${nf.format(row.cohort)}</td><td class="number">${nf.format(row.immature)}</td><td class="number">${nf.format(row.denominator)}</td><td class="number">${nf.format(row.retained)}</td><td class="number">${nf.format(row.notRetained)}</td><td class="rate">${percent(row.retentionRate)}</td><td>${ciText(row)}</td><td class="number">${nf.format(row.manualReview)}</td><td class="number">${nf.format(supplierCount.get(areaKey(row)) || 0)}</td></tr>`
          )
          .join("")
      : '<tr><td colspan="12" class="empty-cell">暂无结果</td></tr>';

    $("riskBlind").textContent = nf.format(summary.blindEliminationRisk);
    $("riskNoAttendance").textContent = nf.format(
      state.results.filter((row) => row.manualFlags.includes("系统派遣覆盖T但窗口无有效出勤")).length
    );
    $("riskLE").textContent = nf.format(state.results.filter((row) => row.latestEndRelation === "L<E").length);
    $("riskCoverage").textContent = state.attendanceStore
      ? `${day(state.attendanceStore.coverageMin)}
${day(state.attendanceStore.coverageMax)}`
      : "—";
  }

  function populateCountryFilters() {
    const countries = Array.from(
      new Map(state.results.map((row) => [row.country, { country: row.country, globalRegion: row.globalRegion }])).values()
    ).sort((a, b) => {
      const region = GLOBAL_REGION_ORDER.indexOf(a.globalRegion) - GLOBAL_REGION_ORDER.indexOf(b.globalRegion);
      return region || a.country.localeCompare(b.country, "zh-CN");
    });
    for (const select of [$("supplierCountryFilter"), $("peopleCountryFilter")]) {
      const current = select.value;
      select.innerHTML = `<option value="">全部国家</option>${countries
        .map((row) => `<option value="${escapeHtml(row.country)}">${escapeHtml(`${row.globalRegion}｜${row.country}`)}</option>`)
        .join("")}`;
      if (countries.some((row) => row.country === current)) select.value = current;
    }
  }

  function compareValues(a, b, key, direction) {
    let left = a[key];
    let right = b[key];
    if (left == null) left = direction === "asc" ? Infinity : -Infinity;
    if (right == null) right = direction === "asc" ? Infinity : -Infinity;
    let result;
    if (typeof left === "number" && typeof right === "number") result = left - right;
    else result = String(left).localeCompare(String(right), "zh-CN", { numeric: true });
    return direction === "asc" ? result : -result;
  }

  function renderSuppliers() {
    if (!state.aggregate) return;
    const country = $("supplierCountryFilter").value;
    const query = $("supplierSearch").value.trim().toLowerCase();
    const sample = $("sampleFilter").value;
    const sort = state.supplierSort;
    const rows = state.aggregate.suppliers
      .filter((row) => !country || row.country === country)
      .filter((row) => !query || `${row.supplierName} ${row.supplierCode}`.toLowerCase().includes(query))
      .filter((row) => !sample || (sample === "comparable" ? row.denominator >= 10 : row.denominator < 10))
      .sort((a, b) => compareValues(a, b, sort.key, sort.direction));
    $("supplierResultCount").textContent = `${nf.format(rows.length)}条`;
    $("supplierTableBody").innerHTML = rows.length
      ? rows
          .map(
            (row) => `<tr><td>${escapeHtml(row.globalRegion)}</td><td>${escapeHtml(row.country)}</td><td>${escapeHtml(row.region)}</td><td title="${escapeHtml(row.supplierName)}">${escapeHtml(row.supplierName || "未识别")}</td><td class="number">${nf.format(row.cohort)}</td><td class="number">${nf.format(row.immature)}</td><td class="number">${nf.format(row.denominator)}</td><td class="number">${nf.format(row.retained)}</td><td class="number">${nf.format(row.notRetained)}</td><td class="rate">${percent(row.retentionRate)}</td><td>${ciText(row)}</td><td class="number">${nf.format(row.manualReview)}</td><td class="number">${nf.format(row.blindEliminationRisk)}</td><td>${escapeHtml(row.sampleFlag)}</td></tr>`
          )
          .join("")
      : '<tr><td colspan="14" class="empty-cell">没有符合条件的供应商</td></tr>';
  }

  function filteredPeople() {
    const country = $("peopleCountryFilter").value;
    const result = $("peopleResultFilter").value;
    const review = $("peopleReviewFilter").value;
    const query = $("peopleSearch").value.trim().toLowerCase();
    return state.results.filter((row) => {
      if (country && row.country !== country) return false;
      if (result && row.finalResult !== result) return false;
      if (review === "yes" && !row.manualReview) return false;
      if (review === "no" && row.manualReview) return false;
      if (query && !`${row.employeeId} ${row.name} ${row.supplierName}`.toLowerCase().includes(query)) return false;
      return true;
    });
  }

  function overrideSelect(row) {
    const options = ["", "留存", "未留存", "未成熟", "待核验", "豁免"];
    const encodedId = encodeURIComponent(row.employeeId);
    return `<select class="review-select" data-override-result="${encodedId}" aria-label="人工调整${escapeHtml(row.employeeId)}">${options
      .map(
        (value) => `<option value="${value}" ${row.overrideResult === value ? "selected" : ""}>${value || "不调整"}</option>`
      )
      .join("")}</select>`;
  }

  function renderPeople() {
    const rows = filteredPeople();
    const totalPages = Math.max(1, Math.ceil(rows.length / state.pageSize));
    state.peoplePage = Math.min(Math.max(1, state.peoplePage), totalPages);
    const start = (state.peoplePage - 1) * state.pageSize;
    const pageRows = rows.slice(start, start + state.pageSize);
    $("peopleResultCount").textContent = `${nf.format(rows.length)}人`;
    $("peoplePageInfo").textContent = rows.length ? `第${state.peoplePage}/${totalPages}页` : "第0/0页";
    $("peoplePrev").disabled = !rows.length || state.peoplePage <= 1;
    $("peopleNext").disabled = !rows.length || state.peoplePage >= totalPages;
    $("peopleTableBody").innerHTML = pageRows.length
      ? pageRows
          .map(
            (row) => `<tr><td>${escapeHtml(row.globalRegion)}</td><td>${escapeHtml(row.country)}</td><td>${escapeHtml(row.region)}</td><td title="${escapeHtml(row.supplierName)}">${escapeHtml(row.supplierName)}</td><td>${escapeHtml(row.name)}</td><td>${escapeHtml(row.employeeId)}</td><td>${day(row.firstDispatch)}</td><td>${day(row.observationDay)}</td><td>${day(row.windowStart)}~${day(row.observationDay)}</td><td>${day(row.latestDispatch) || "—"}</td><td>${day(row.latestEnd) || "—"}</td><td class="number">${row.validWindowDays.length}</td><td>${statusBadge(row.modelResult)}</td><td title="${escapeHtml(uiText(row.reason))}">${escapeHtml(uiText(row.reason))}</td><td>${row.manualReview ? escapeHtml(uiText(row.manualFlags.join("；"))) : "否"}</td><td>${overrideSelect(row)}</td></tr>`
          )
          .join("")
      : '<tr><td colspan="16" class="empty-cell">没有符合条件的人员</td></tr>';
  }

  function suggestedAction(row) {
    if (row.manualFlags.includes("交叉表L/E与主表不一致")) return "确认两份导出的时点和最新性，以可追溯的最新L/E重新计算";
    if (row.manualFlags.includes("L日之后仍有有效出勤")) return "核对L的业务含义与考勤原始行；有效出勤不应被过期L日静默覆盖";
    if (row.manualFlags.includes("L<E且T落在冲突区间")) return "回溯OTWS派遣删除/结束记录，确认T日真实派遣状态";
    if (row.manualFlags.includes("T窗口供应商与入职供应商不一致")) return "核实是否雇主转换，决定原供应商归责与是否豁免";
    if (row.manualFlags.includes("窗口存在未确认正工时")) return "先完成考勤确认/复核，再重新计算";
    if (row.manualFlags.includes("系统派遣覆盖T但窗口无有效出勤")) return "核实批准休假、区域未排班或系统漏结束；无证据则维持未留存";
    return "核对T边界日与原始派遣/考勤凭证";
  }

  function renderReview() {
    const rows = state.results.filter((row) => row.manualReview);
    $("reviewTableBody").innerHTML = rows.length
      ? rows
          .map((row) => {
            const encodedId = encodeURIComponent(row.employeeId);
            return `<tr><td>${escapeHtml(row.globalRegion)}</td><td>${escapeHtml(row.country)}</td><td>${escapeHtml(row.region)}</td><td title="${escapeHtml(row.supplierName)}">${escapeHtml(row.supplierName)}</td><td>${escapeHtml(row.name)}<br><small>${escapeHtml(row.employeeId)}</small></td><td>${day(row.observationDay)}</td><td>最晚 ${day(row.latestDispatch) || "—"}<br>结束 ${day(row.latestEnd) || "—"}</td><td>${row.validWindowDays.length}天<br><small>${escapeHtml(row.validWindowDays.map(day).join("、"))}</small></td><td>${statusBadge(row.modelResult)}<br><small>${escapeHtml(uiText(row.reason))}</small></td><td title="${escapeHtml(uiText(suggestedAction(row)))}">${escapeHtml(uiText(row.manualFlags.join("；")))}</td><td>${overrideSelect(row)}</td><td><input class="review-note" data-override-note="${encodedId}" value="${escapeHtml(row.overrideNote)}" placeholder="证据或调整说明" /></td></tr>`;
          })
          .join("")
      : '<tr><td colspan="12" class="empty-cell">暂无需复核记录</td></tr>';
  }

  function updateOverride(employeeId, field, value) {
    const current = state.overrides.get(employeeId) || { result: "", note: "" };
    current[field] = value;
    if (!current.result && !current.note) state.overrides.delete(employeeId);
    else state.overrides.set(employeeId, current);
  }

  function handleOverrideChange(event) {
    const resultSelect = event.target.closest("[data-override-result]");
    if (resultSelect) {
      const employeeId = decodeURIComponent(resultSelect.dataset.overrideResult);
      updateOverride(employeeId, "result", resultSelect.value);
      calculate();
      return;
    }
    const noteInput = event.target.closest("[data-override-note]");
    if (noteInput) {
      const employeeId = decodeURIComponent(noteInput.dataset.overrideNote);
      updateOverride(employeeId, "note", noteInput.value.trim());
      const row = state.results.find((item) => item.employeeId === employeeId);
      if (row) row.overrideNote = noteInput.value.trim();
    }
  }

  async function exportWorkbook() {
    if (!state.aggregate || !state.results.length || !ReportExporter) return;
    clearDownloadArtifact();
    dom.exportBtn.disabled = true;
    dom.exportBtn.textContent = "正在生成报表...";
    setMessage("正在按参考版模板生成七张工作表，请稍候。", "neutral");
    try {
      const report = await ReportExporter.buildReport({
        aggregate: state.aggregate,
        results: state.results,
        profiles: state.profiles,
        attendanceStore: state.attendanceStore,
        periodDays: state.periodDays,
        windowDays: state.windowDays,
        startLabel: dom.cohortStart.value,
        endLabel: dom.cohortEnd.value,
        scopeLabel: currentScopeLabel(),
        dayToISO: Engine.dayToISO,
      });
      exposeDownloadArtifact(report.blob, report.fileName);
      setMessage("报表已生成并尝试自动下载。若浏览器没有弹出下载，请点击顶部【下载已生成报表】。", "success");
    } catch (error) {
      console.error(error);
      setMessage(`报表生成失败：${error.message || error}`, "error");
    } finally {
      dom.exportBtn.disabled = false;
      dom.exportBtn.textContent = "导出留存率报表";
    }
  }

  function reset() {
    clearDownloadArtifact();
    for (const input of [dom.cohortFile, dom.latestFile, dom.attendanceFiles]) input.value = "";
    updateFileLabel(dom.cohortFile, dom.cohortFileName, false);
    updateFileLabel(dom.latestFile, dom.latestFileName, false);
    updateFileLabel(dom.attendanceFiles, dom.attendanceFileName, true);
    state.cohortMap = null;
    state.latestMap = new Map();
    state.attendanceStore = null;
    state.allResults = [];
    state.results = [];
    state.aggregate = null;
    state.scope = "";
    state.scopeInitialized = false;
    state.profiles = { cohort: null, latest: null, attendance: [] };
    state.overrides.clear();
    dom.cohortStart.value = "";
    dom.cohortEnd.value = "";
    dom.scopeSelect.innerHTML = '<option value="">全部已识别大区</option>';
    dom.scopeSelect.disabled = true;
    setProgress(0, "等待上传数据");
    setMessage("已清空本地数据。上传数据后，程序会先检查字段、日期覆盖和重复工号。", "neutral");
    dom.calculateBtn.disabled = true;
    dom.exportBtn.disabled = true;
    updateCoverageHint();
    window.location.hash = "";
    window.location.reload();
  }

  function bindEvents() {
    window.addEventListener("beforeunload", clearDownloadArtifact);
    dom.cohortFile.addEventListener("change", () => updateFileLabel(dom.cohortFile, dom.cohortFileName, false));
    dom.latestFile.addEventListener("change", () => updateFileLabel(dom.latestFile, dom.latestFileName, false));
    dom.attendanceFiles.addEventListener("change", () => updateFileLabel(dom.attendanceFiles, dom.attendanceFileName, true));
    dom.loadBtn.addEventListener("click", loadData);
    dom.calculateBtn.addEventListener("click", calculate);
    dom.exportBtn.addEventListener("click", exportWorkbook);
    dom.resetBtn.addEventListener("click", reset);
    dom.cohortStart.addEventListener("change", updateCoverageHint);
    dom.cohortEnd.addEventListener("change", updateCoverageHint);
    dom.scopeSelect.addEventListener("change", () => {
      state.scope = dom.scopeSelect.value;
      clearDownloadArtifact();
      applyScopeResults();
      setMessage(`已切换至${currentScopeLabel()}统计范围；逐人判定口径未改变。`, "success");
    });

    configureSegments("periodSegments", dom.periodCustomWrap, dom.periodCustom);
    configureSegments("windowSegments", dom.windowCustomWrap, dom.windowCustom);

    for (const tab of document.querySelectorAll(".tabs button[data-tab]")) {
      tab.addEventListener("click", () => {
        for (const button of document.querySelectorAll(".tabs button")) button.classList.toggle("active", button === tab);
        for (const panel of document.querySelectorAll(".tab-panel")) panel.classList.toggle("active", panel.dataset.panel === tab.dataset.tab);
      });
    }

    for (const control of document.querySelectorAll("[data-drop-target]")) {
      const input = $(control.dataset.dropTarget);
      control.addEventListener("dragover", (event) => {
        event.preventDefault();
        control.classList.add("dragging");
      });
      control.addEventListener("dragleave", () => control.classList.remove("dragging"));
      control.addEventListener("drop", (event) => {
        event.preventDefault();
        control.classList.remove("dragging");
        const transfer = new DataTransfer();
        const files = Array.from(event.dataTransfer.files || []).filter((file) => /\.(xlsx|xls|csv)$/i.test(file.name));
        for (const file of input.multiple ? files : files.slice(0, 1)) transfer.items.add(file);
        input.files = transfer.files;
        input.dispatchEvent(new Event("change", { bubbles: true }));
      });
    }

    for (const id of ["supplierCountryFilter", "supplierSearch", "sampleFilter"]) {
      $(id).addEventListener(id === "supplierSearch" ? "input" : "change", renderSuppliers);
    }
    for (const id of ["peopleCountryFilter", "peopleResultFilter", "peopleReviewFilter", "peopleSearch"]) {
      $(id).addEventListener(id === "peopleSearch" ? "input" : "change", () => {
        state.peoplePage = 1;
        renderPeople();
      });
    }
    $("peoplePrev").addEventListener("click", () => {
      state.peoplePage -= 1;
      renderPeople();
    });
    $("peopleNext").addEventListener("click", () => {
      state.peoplePage += 1;
      renderPeople();
    });
    $("supplierTable").addEventListener("click", (event) => {
      const header = event.target.closest("th[data-sort]");
      if (!header) return;
      const key = header.dataset.sort;
      if (state.supplierSort.key === key) {
        state.supplierSort.direction = state.supplierSort.direction === "asc" ? "desc" : "asc";
      } else {
        state.supplierSort = {
          key,
          direction: ["globalRegion", "country", "region", "supplierName"].includes(key) ? "asc" : "desc",
        };
      }
      renderSuppliers();
    });
    $("peopleTableBody").addEventListener("change", handleOverrideChange);
    $("reviewTableBody").addEventListener("change", handleOverrideChange);
  }

  function initialize() {
    if (XLSX && Engine && ReportExporter && window.XlsxPopulate && window.RETENTION_REPORT_TEMPLATE_BASE64) {
      $("libraryStatus").textContent = `Excel解析与模板导出模块 ${XLSX.version || "local"} 已就绪`;
    } else {
      $("libraryStatus").textContent = "Excel解析或模板导出模块加载失败";
      setMessage("缺少本地Excel解析或报表模板文件，请检查vendor目录。", "error");
    }
    bindEvents();
    syncParameters();
  }

  initialize();
})();

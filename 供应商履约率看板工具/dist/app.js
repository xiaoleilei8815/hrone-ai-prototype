(function () {
  "use strict";

  const engine = window.SupplierFulfillmentEngine;
  const exporter = window.SupplierFulfillmentReportExporter;

  if (!engine || !exporter) {
    throw new Error("履约率计算或报表导出引擎未加载");
  }

  const state = {
    files: [],
    analysis: null,
    supplierMetric: "raw",
    activeTab: "overview",
    resizeTimer: null,
    downloadUrl: null,
    downloadFileName: null,
  };

  const palette = {
    navy: "#075985",
    teal: "#0f766e",
    orange: "#d97706",
    gray: "#64748b",
    red: "#dc2626",
    green: "#15803d",
    grid: "#dbe3e8",
    ink: "#334155",
    muted: "#718096",
    region: ["#075985", "#0f766e", "#d97706", "#7c3aed", "#dc2626", "#0891b2"],
  };

  const $ = (id) => document.getElementById(id);

  const elements = {
    fileInput: $("fileInput"),
    dropzone: $("dropzone"),
    fileSummary: $("fileSummary"),
    fileCount: $("fileCount"),
    fileSize: $("fileSize"),
    fileList: $("fileList"),
    clearFilesButton: $("clearFilesButton"),
    analyzeButton: $("analyzeButton"),
    exportButton: $("exportButton"),
    statusPanel: $("statusPanel"),
    statusText: $("statusText"),
    emptyState: $("emptyState"),
    dashboard: $("dashboard"),
    marketFilter: $("marketFilter"),
    regionFilter: $("regionFilter"),
    jobFilter: $("jobFilter"),
    supplierSearch: $("supplierSearch"),
    toast: $("toast"),
    downloadLink: $("downloadLink"),
  };

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function formatBytes(bytes) {
    const value = Number(bytes) || 0;
    if (value < 1024) return `${value} B`;
    if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`;
    return `${(value / 1024 / 1024).toFixed(1)} MB`;
  }

  function formatNumber(value, maximumFractionDigits = 2) {
    if (value == null || !Number.isFinite(Number(value))) return "—";
    return Number(value).toLocaleString("zh-CN", {
      minimumFractionDigits: 0,
      maximumFractionDigits,
    });
  }

  function formatRate(value, digits = 1) {
    if (value == null || !Number.isFinite(Number(value))) return "—";
    return `${(Number(value) * 100).toFixed(digits)}%`;
  }

  function formatSignedRate(value, digits = 1) {
    if (value == null || !Number.isFinite(Number(value))) return "—";
    const points = Number(value) * 100;
    return `${points > 0 ? "+" : ""}${points.toFixed(digits)}pp`;
  }

  function rateClass(value) {
    if (value == null) return "empty-cell";
    if (value >= 0.9) return "good";
    if (value >= 0.75) return "warn";
    return "bad";
  }

  function sum(rows, field) {
    return rows.reduce((total, row) => total + (Number(row[field]) || 0), 0);
  }

  function mean(values) {
    const valid = values.filter((value) => value != null && Number.isFinite(value));
    return valid.length ? valid.reduce((total, value) => total + value, 0) / valid.length : null;
  }

  function setStatus(text, type) {
    elements.statusText.textContent = text;
    elements.statusPanel.classList.remove("is-working", "is-error", "is-success");
    if (type) elements.statusPanel.classList.add(`is-${type}`);
  }

  function showToast(message, duration = 4200) {
    elements.toast.textContent = message;
    elements.toast.hidden = false;
    window.clearTimeout(showToast.timer);
    showToast.timer = window.setTimeout(() => {
      elements.toast.hidden = true;
    }, duration);
  }

  function setBusy(isBusy, label) {
    elements.analyzeButton.disabled = isBusy || state.files.length === 0;
    elements.exportButton.disabled = isBusy || !state.analysis;
    elements.clearFilesButton.disabled = isBusy || state.files.length === 0;
    if (label) setStatus(label, "working");
  }

  function clearReadyDownload() {
    if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
    state.downloadUrl = null;
    state.downloadFileName = null;
    elements.downloadLink.removeAttribute("href");
    elements.downloadLink.removeAttribute("download");
    elements.downloadLink.hidden = true;
    elements.exportButton.innerHTML = '<span class="button-icon" aria-hidden="true">⇩</span>导出报表';
  }

  function prepareDownload(blob, fileName) {
    clearReadyDownload();
    state.downloadUrl = URL.createObjectURL(blob);
    state.downloadFileName = fileName;
    elements.downloadLink.href = state.downloadUrl;
    elements.downloadLink.download = fileName;
    elements.downloadLink.textContent = `下载 ${fileName}`;
    elements.downloadLink.hidden = false;
    elements.exportButton.innerHTML = '<span class="button-icon" aria-hidden="true">⇩</span>再次点击下载';
  }

  function triggerReadyDownload() {
    if (!state.downloadUrl) return false;
    elements.downloadLink.click();
    setStatus(`报表已生成：${state.downloadFileName}`, "success");
    return true;
  }

  function fileIdentity(file) {
    return `${file.name}\u0001${file.size}\u0001${file.lastModified}`;
  }

  function addFiles(fileList) {
    const existing = new Set(state.files.map(fileIdentity));
    Array.from(fileList || []).forEach((file) => {
      if (!/\.(xlsx|xls)$/i.test(file.name)) return;
      const identity = fileIdentity(file);
      if (!existing.has(identity)) {
        existing.add(identity);
        state.files.push(file);
      }
    });
    state.analysis = null;
    clearReadyDownload();
    renderFileList();
    resetAnalysisPresentation();
  }

  function renderFileList() {
    const totalBytes = state.files.reduce((total, file) => total + file.size, 0);
    elements.fileSummary.hidden = state.files.length === 0;
    elements.fileCount.textContent = `${state.files.length} 个文件`;
    elements.fileSize.textContent = formatBytes(totalBytes);
    elements.fileList.innerHTML = state.files
      .map(
        (file, index) => `
          <div class="file-item">
            <span class="file-item-icon" aria-hidden="true">X</span>
            <span class="file-item-name" title="${escapeHtml(file.name)}">${escapeHtml(
              file.name,
            )}</span>
            <button class="icon-button" data-remove-file="${index}" type="button"
              title="移除${escapeHtml(file.name)}" aria-label="移除${escapeHtml(
                file.name,
              )}">×</button>
          </div>`,
      )
      .join("");
    elements.analyzeButton.disabled = state.files.length === 0;
    elements.clearFilesButton.disabled = state.files.length === 0;
    setStatus(
      state.files.length ? `已选择 ${state.files.length} 个文件，等待计算` : "等待上传源文件",
      null,
    );
  }

  function clearFiles() {
    state.files = [];
    state.analysis = null;
    clearReadyDownload();
    elements.fileInput.value = "";
    renderFileList();
    resetAnalysisPresentation();
  }

  function resetAnalysisPresentation() {
    elements.exportButton.disabled = true;
    elements.emptyState.hidden = false;
    elements.dashboard.hidden = true;
    $("scopeRegion").textContent = "待识别";
    $("scopeMarket").textContent = "待识别";
    $("scopePeriod").textContent = "待识别";
    $("scopeSupplier").textContent = "待识别";
    $("scopeCoverage").textContent = "待识别";
    $("qualityBadge").textContent = "WAIT";
    $("qualityBadge").className = "quality-badge neutral";
    $("qualityMiniList").innerHTML = `
      <div class="quality-mini-row"><span>三表月×区域</span><span>—</span></div>
      <div class="quality-mini-row"><span>原子发单/派遣</span><span>—</span></div>
      <div class="quality-mini-row"><span>责任闭合</span><span>—</span></div>`;
  }

  async function analyze() {
    if (!state.files.length) return;
    clearReadyDownload();
    setBusy(true, "正在逐日展开区间日期并构建原子责任底表…");
    await new Promise((resolve) => window.requestAnimationFrame(resolve));
    try {
      state.analysis = await engine.analyzeFiles(state.files);
      populateFilters();
      renderAnalysis();
      setStatus(
        `计算完成：${formatNumber(state.analysis.quality.clean_row_counts.unit_atomic, 0)} 个仓库原子单元，${formatNumber(
          state.analysis.quality.clean_row_counts.supplier_atomic,
          0,
        )} 条供应商责任记录`,
        "success",
      );
      showToast(`${state.analysis.meta.title}已通过计算链路`);
    } catch (error) {
      console.error(error);
      state.analysis = null;
      resetAnalysisPresentation();
      setStatus(error.message || "统计失败，请检查文件结构", "error");
      showToast(error.message || "统计失败，请检查文件结构", 7000);
    } finally {
      setBusy(false);
    }
  }

  function populateFilters() {
    const analysis = state.analysis;
    const markets = analysis.meta.markets || [];
    elements.marketFilter.innerHTML =
      markets.length > 1
        ? [
            '<option value="__all__">全部大区</option>',
            ...markets.map(
              (market) =>
                `<option value="${escapeHtml(market)}">${escapeHtml(market)}</option>`,
            ),
          ].join("")
        : `<option value="__all__">${escapeHtml(markets[0] || "全部大区")}</option>`;
    elements.marketFilter.value = "__all__";
    populateRegionFilter();
    const jobOptions = [
      '<option value="__all__">全部工种</option>',
      ...analysis.meta.jobs.map(
        (job) => `<option value="${escapeHtml(job)}">${escapeHtml(job)}</option>`,
      ),
    ];
    elements.jobFilter.innerHTML = jobOptions.join("");
    elements.jobFilter.value = "__all__";
    elements.supplierSearch.value = "";
  }

  function populateRegionFilter() {
    const market = elements.marketFilter.value || "__all__";
    const regions = state.analysis.meta.regions.filter(
      (region) =>
        market === "__all__" || state.analysis.meta.region_market[region] === market,
    );
    elements.regionFilter.innerHTML = [
      `<option value="__all__">${regions.length > 1 ? "全部运营区域" : regions[0]}</option>`,
      ...(regions.length > 1
        ? regions.map(
            (region) =>
              `<option value="${escapeHtml(region)}">${escapeHtml(region)}</option>`,
          )
        : []),
    ].join("");
    elements.regionFilter.value = "__all__";
  }

  function selection() {
    return {
      market: elements.marketFilter.value || "__all__",
      region: elements.regionFilter.value || "__all__",
      job: elements.jobFilter.value || "__all__",
    };
  }

  function contextTarget() {
    const analysis = state.analysis;
    const { market, region, job } = selection();
    const suffix = job === "__all__" ? "total" : "job";
    if (region !== "__all__") {
      return { rowType: `region_${suffix}`, targetRegion: region };
    }
    if (market !== "__all__" && analysis.meta.markets.length > 1) {
      return { rowType: `market_${suffix}`, targetRegion: `${market}总计` };
    }
    if (analysis.meta.regions.length > 1) {
      return { rowType: `scope_${suffix}`, targetRegion: analysis.meta.scope_total_label };
    }
    return { rowType: `region_${suffix}`, targetRegion: analysis.meta.regions[0] };
  }

  function contextDisplayRow() {
    const analysis = state.analysis;
    const { job } = selection();
    const { rowType, targetRegion } = contextTarget();
    return analysis.display_metrics.find(
      (row) =>
        row.row_type === rowType &&
        row.region === targetRegion &&
        (job === "__all__" || row.job === job),
    );
  }

  function contextWarehouseRows() {
    const analysis = state.analysis;
    const { job } = selection();
    const { rowType, targetRegion } = contextTarget();
    return analysis.warehouse_display.filter(
      (row) =>
        row.row_type === rowType &&
        row.region === targetRegion &&
        (job === "__all__" || row.job === job),
    );
  }

  function contextLabel() {
    const { market, region, job } = selection();
    const regionLabel =
      region !== "__all__"
        ? region.replace(/区$/, "")
        : market !== "__all__"
          ? market
          : state.analysis.meta.scope_prefix;
    return job === "__all__" ? `${regionLabel}·全部工种` : `${regionLabel}·${job}`;
  }

  function renderAnalysis() {
    const analysis = state.analysis;
    elements.emptyState.hidden = true;
    elements.dashboard.hidden = false;
    elements.exportButton.disabled = false;
    $("reportEyebrow").textContent = `${analysis.meta.primary_metric} · 主指标`;
    $("reportTitle").textContent = analysis.meta.title;
    $("reportMeta").textContent = `${analysis.meta.source_file_count} 个源文件 · ${analysis.meta.markets.length} 个大区 · ${analysis.meta.regions.length} 个运营区域 · ${analysis.meta.supplier_count} 家供应商 · 单位：人日`;
    renderScope();
    renderQualitySidebar();
    renderAllViews();
    window.requestAnimationFrame(() => {
      drawOverviewCharts();
      if (window.matchMedia("(max-width: 860px)").matches) {
        elements.dashboard.scrollIntoView({ block: "start" });
      }
    });
  }

  function renderScope() {
    const { meta } = state.analysis;
    $("scopeMarket").textContent = meta.markets.join("、");
    $("scopeRegion").textContent = meta.regions.join("、");
    $("scopePeriod").textContent = `${meta.period_start} 至 ${meta.period_end}`;
    $("scopeSupplier").textContent = `${meta.supplier_count} 家`;
    $("scopeCoverage").textContent = `${meta.coverage.covered_days}/${meta.coverage.expected_days} 天`;
  }

  function renderQualitySidebar() {
    const quality = state.analysis.quality;
    const badge = $("qualityBadge");
    badge.textContent = quality.status;
    badge.className = `quality-badge ${quality.status === "PASS" ? "pass" : "review"}`;
    const ids = [
      ["three_way_month_region", "三表月×区域"],
      ["demand_dispatch_atomic", "原子发单/派遣"],
      ["responsibility_closure", "责任闭合"],
    ];
    $("qualityMiniList").innerHTML = ids
      .map(([id, label]) => {
        const check = quality.critical_checks.find((item) => item.id === id);
        const passed = check?.status === "PASS";
        return `<div class="quality-mini-row"><span>${label}</span><strong class="${
          passed ? "" : "fail"
        }">${passed ? "PASS" : `REVIEW ${check?.difference_count || 0}`}</strong></div>`;
      })
      .join("");
  }

  function renderAllViews() {
    renderOverview();
    renderSupplierTable();
    renderWarehouseView();
    renderQualityView();
    renderMethodView();
  }

  function renderOverview() {
    renderKpis();
    renderTrendLegend();
    renderAllocation();
    renderRegionBars();
    renderSupplierPreview();
  }

  function renderKpis() {
    const display = contextDisplayRow();
    const warehouseRows = contextWarehouseRows();
    const warehouseAverage = mean(
      warehouseRows.map((row) => row.warehouse_fulfillment_rate),
    );
    const coverageAverage = mean(warehouseRows.map((row) => row.issue_coverage_rate));
    const warehouseWeighted = engine.safeRate(
      sum(warehouseRows, "warehouse_effective"),
      sum(warehouseRows, "true_demand"),
    );
    const coverageWeighted = engine.safeRate(
      sum(warehouseRows, "issue_covered"),
      sum(warehouseRows, "true_demand"),
    );
    const cards = [
      {
        label: "原始接单兑现率（月均）",
        value: formatRate(display?.raw_average),
        note: `统计期加权 ${formatRate(display?.raw_weighted)} · 主指标`,
        className: "",
      },
      {
        label: "超发校正履约率（月均）",
        value: formatRate(display?.adjusted_average),
        note: `统计期加权 ${formatRate(display?.adjusted_weighted)} · 对照指标`,
        className: "accent-orange",
      },
      {
        label: "仓库需求满足率（月均）",
        value: formatRate(warehouseAverage),
        note: `统计期加权 ${formatRate(warehouseWeighted)} · 仓库保障`,
        className: "accent-teal",
      },
      {
        label: "发单覆盖率（月均）",
        value: formatRate(coverageAverage),
        note: `统计期加权 ${formatRate(coverageWeighted)} · HR配置`,
        className: coverageAverage != null && coverageAverage < 0.9 ? "accent-red" : "",
      },
    ];
    $("kpiGrid").innerHTML = cards
      .map(
        (card) => `
          <article class="kpi-card ${card.className}">
            <span class="kpi-label">${card.label}</span>
            <strong class="kpi-value">${card.value}</strong>
            <span class="kpi-note">${card.note}</span>
          </article>`,
      )
      .join("");
  }

  function renderTrendLegend() {
    const legend = [
      ["原始兑现", palette.navy],
      ["超发校正", palette.orange],
      ["仓库满足", palette.teal],
      ["发单覆盖", palette.gray],
    ];
    $("trendLegend").innerHTML = legend
      .map(
        ([label, color]) =>
          `<span class="legend-item"><i class="legend-swatch" style="background:${color}"></i>${label}</span>`,
      )
      .join("");
  }

  function renderAllocation() {
    const monthMap = new Map(contextWarehouseRows().map((row) => [row.month, row]));
    $("allocationChart").innerHTML = state.analysis.meta.months
      .map((month) => {
        const row = monthMap.get(month.key);
        if (!row) {
          return `<div class="allocation-row"><div class="allocation-label"><strong>${month.label}</strong><span>无有效数据</span></div></div>`;
        }
        const covered = Math.max(0, row.true_demand - row.under_issue);
        const under = Math.max(0, row.under_issue);
        const over = Math.max(0, row.over_issue);
        const scale = Math.max(row.true_demand + over, 1);
        return `
          <div class="allocation-row">
            <div class="allocation-label">
              <strong>${month.label}</strong>
              <span>真实需求 ${formatNumber(row.true_demand)} 人日</span>
            </div>
            <div class="allocation-track" title="覆盖 ${formatNumber(
              covered,
            )}；少发 ${formatNumber(under)}；超发 ${formatNumber(over)}">
              <span class="allocation-segment covered" style="width:${(covered / scale) * 100}%"></span>
              <span class="allocation-segment under" style="width:${(under / scale) * 100}%"></span>
              <span class="allocation-segment over" style="width:${(over / scale) * 100}%"></span>
            </div>
            <div class="allocation-stats">
              <span>覆盖 ${formatNumber(covered)}</span>
              <span>少发 ${formatNumber(under)}</span>
              <span>超发 ${formatNumber(over)}</span>
            </div>
          </div>`;
      })
      .join("");
  }

  function regionRowsForCurrentJob() {
    const { market, region, job } = selection();
    return state.analysis.meta.regions
      .filter(
        (item) =>
          (market === "__all__" || state.analysis.meta.region_market[item] === market) &&
          (region === "__all__" || item === region),
      )
      .map((item) => {
        const row = state.analysis.display_metrics.find(
          (candidate) =>
            candidate.region === item &&
            candidate.row_type === (job === "__all__" ? "region_total" : "region_job") &&
            (job === "__all__" || candidate.job === job),
        );
        return row
          ? {
              region: item,
              raw_average: row.raw_average,
              adjusted_average: row.adjusted_average,
            }
          : null;
      })
      .filter(Boolean);
  }

  function renderRegionBars() {
    const rows = regionRowsForCurrentJob();
    if (!rows.length) {
      $("regionBars").innerHTML = '<p class="empty-cell">当前筛选无区域数据</p>';
      return;
    }
    $("regionBars").innerHTML = rows
      .map((row) => {
        const rawWidth = Math.max(0, Math.min(100, (row.raw_average || 0) * 100));
        const adjustedWidth = Math.max(
          0,
          Math.min(100, (row.adjusted_average || 0) * 100),
        );
        return `
          <div class="region-bar-row">
            <span class="region-bar-label" title="${escapeHtml(row.region)}">${escapeHtml(
              row.region,
            )}</span>
            <div class="region-bar-track">
              <span class="region-bar raw" style="width:${rawWidth}%"></span>
              <span class="region-bar adjusted" style="width:${adjustedWidth}%"></span>
            </div>
            <span class="region-bar-value">${formatRate(row.raw_average, 0)}<br>${formatRate(
              row.adjusted_average,
              0,
            )}</span>
          </div>`;
      })
      .join("");
  }

  function filteredSupplierRows() {
    const { market, region, job } = selection();
    const targetType = job === "__all__" ? "supplier_total" : "supplier_job";
    const search = elements.supplierSearch.value.trim().toLocaleLowerCase("zh-CN");
    return state.analysis.display_metrics.filter((row) => {
      if (row.row_type !== targetType) return false;
      if (market !== "__all__" && row.market !== market) return false;
      if (region !== "__all__" && row.region !== region) return false;
      if (job !== "__all__" && row.job !== job) return false;
      if (
        search &&
        !`${row.supplier_name} ${row.supplier_id}`.toLocaleLowerCase("zh-CN").includes(search)
      ) {
        return false;
      }
      return true;
    });
  }

  function renderSupplierPreview() {
    const rows = filteredSupplierRows()
      .slice()
      .sort(
        (left, right) =>
          (right.raw_average ?? -1) - (left.raw_average ?? -1) ||
          right.totals.issued - left.totals.issued,
      );
    $("supplierPreviewMeta").textContent = `${contextLabel()} · ${rows.length} 家供应商，按原始月均排序`;
    const shown = rows.slice(0, 12);
    $("supplierPreviewBody").innerHTML = shown.length
      ? shown
          .map((row) => {
            const lift =
              row.raw_average != null && row.adjusted_average != null
                ? row.adjusted_average - row.raw_average
                : null;
            return `<tr>
              <td>${escapeHtml(row.region)}</td>
              <td>${escapeHtml(row.supplier_name)}</td>
              <td class="rate-cell ${rateClass(row.raw_average)}">${formatRate(
                row.raw_average,
              )}</td>
              <td class="rate-cell ${rateClass(row.adjusted_average)}">${formatRate(
                row.adjusted_average,
              )}</td>
              <td class="${lift != null && lift >= 0 ? "delta-positive" : "delta-negative"}">${formatSignedRate(
                lift,
              )}</td>
              <td>${formatNumber(row.totals.issued)}</td>
              <td>${formatNumber(row.totals.over_issue_relief)}</td>
            </tr>`;
          })
          .join("")
      : '<tr><td colspan="7" class="empty-cell">当前筛选无供应商数据</td></tr>';
  }

  function renderSupplierTable() {
    if (!state.analysis) return;
    const metric = state.supplierMetric;
    const months = state.analysis.meta.months;
    const denominatorLabel = metric === "raw" ? "累计发单" : "合理责任";
    const numeratorLabel = metric === "raw" ? "有效兑现" : "校正兑现";
    $("supplierTableHead").innerHTML = `<tr>
      <th>区域</th><th>供应商</th><th>供应商ID</th><th>工种</th>
      ${months.map((month) => `<th>${escapeHtml(month.label)}</th>`).join("")}
      <th>月均</th><th>统计期加权</th><th>${denominatorLabel}</th><th>${numeratorLabel}</th>
      ${metric === "adjusted" ? "<th>释放责任</th>" : ""}
    </tr>`;
    const rows = filteredSupplierRows()
      .slice()
      .sort(
        (left, right) =>
          left.region.localeCompare(right.region, "zh-CN") ||
          (right[`${metric}_average`] ?? -1) - (left[`${metric}_average`] ?? -1) ||
          left.supplier_name.localeCompare(right.supplier_name, "zh-CN"),
      );
    $("supplierTableBody").innerHTML = rows.length
      ? rows
          .map((row) => {
            const monthCells = months
              .map((month) => {
                const value = row.by_month.find((item) => item.month === month.key)?.[
                  `${metric}_rate`
                ];
                return `<td class="rate-cell ${rateClass(value)}">${formatRate(value)}</td>`;
              })
              .join("");
            const average = row[`${metric}_average`];
            const weighted = row[`${metric}_weighted`];
            const denominator =
              metric === "raw" ? row.totals.issued : row.totals.reasonable_responsibility;
            const numerator =
              metric === "raw" ? row.totals.raw_effective : row.totals.adjusted_effective;
            return `<tr>
              <td>${escapeHtml(row.region)}</td>
              <td>${escapeHtml(row.supplier_name)}</td>
              <td>${escapeHtml(row.supplier_id)}</td>
              <td>${escapeHtml(row.job)}</td>
              ${monthCells}
              <td class="rate-cell ${rateClass(average)}">${formatRate(average)}</td>
              <td class="rate-cell ${rateClass(weighted)}">${formatRate(weighted)}</td>
              <td>${formatNumber(denominator)}</td>
              <td>${formatNumber(numerator)}</td>
              ${metric === "adjusted" ? `<td>${formatNumber(row.totals.over_issue_relief)}</td>` : ""}
            </tr>`;
          })
          .join("")
      : `<tr><td colspan="${months.length + (metric === "adjusted" ? 10 : 9)}" class="empty-cell">当前筛选无供应商数据</td></tr>`;
  }

  function warehouseTableRows() {
    const { market, region, job } = selection();
    return state.analysis.warehouse_display.filter((row) => {
      if (job !== "__all__" && row.job !== job) return false;
      if (region !== "__all__") {
        if (row.region !== region) return false;
        if (job !== "__all__") return row.row_type === "region_job";
        return row.row_type === "region_job" || row.row_type === "region_total";
      }
      if (market !== "__all__" && state.analysis.meta.markets.length > 1) {
        const inMarket = row.market === market;
        if (!inMarket) return false;
        return job === "__all__"
          ? ["region_total", "market_total"].includes(row.row_type)
          : ["region_job", "market_job"].includes(row.row_type);
      }
      return job === "__all__"
        ? ["region_total", "market_total", "scope_total"].includes(row.row_type)
        : ["region_job", "market_job", "scope_job"].includes(row.row_type);
    });
  }

  function renderWarehouseView() {
    if (!state.analysis) return;
    const contextRows = contextWarehouseRows();
    const trueDemand = sum(contextRows, "true_demand");
    const issued = sum(contextRows, "issued");
    const overIssue = sum(contextRows, "over_issue");
    const underIssue = sum(contextRows, "under_issue");
    const cards = [
      ["统计期真实需求", `${formatNumber(trueDemand)} 人日`, "基础需求 + 有效补员", ""],
      ["统计期总发单", `${formatNumber(issued)} 人日`, `发单/需求 ${formatRate(engine.safeRate(issued, trueDemand))}`, ""],
      ["累计超发", `${formatNumber(overIssue)} 人日`, "按原子单元逐日判定", "accent-orange"],
      ["累计少发", `${formatNumber(underIssue)} 人日`, "真实需求未被发单覆盖", "accent-red"],
    ];
    $("warehouseKpis").innerHTML = cards
      .map(
        ([label, value, note, className]) => `<article class="kpi-card ${className}">
          <span class="kpi-label">${label}</span><strong class="kpi-value">${value}</strong>
          <span class="kpi-note">${note}</span></article>`,
      )
      .join("");

    const rows = warehouseTableRows();
    $("warehouseTableBody").innerHTML = rows.length
      ? rows
          .map(
            (row) => `<tr>
              <td>${escapeHtml(row.region)}</td>
              <td>${escapeHtml(row.level)}</td>
              <td>${escapeHtml(row.job)}</td>
              <td>${escapeHtml(row.month)}</td>
              <td>${formatNumber(row.true_demand)}</td>
              <td>${formatNumber(row.issued)}</td>
              <td>${formatNumber(row.dispatched)}</td>
              <td class="rate-cell ${rateClass(row.issue_coverage_rate)}">${formatRate(
                row.issue_coverage_rate,
              )}</td>
              <td class="rate-cell ${rateClass(row.warehouse_fulfillment_rate)}">${formatRate(
                row.warehouse_fulfillment_rate,
              )}</td>
              <td>${formatNumber(row.over_issue)}</td>
              <td>${formatNumber(row.under_issue)}</td>
            </tr>`,
          )
          .join("")
      : '<tr><td colspan="11" class="empty-cell">当前筛选无仓库数据</td></tr>';
  }

  function renderQualityView() {
    if (!state.analysis) return;
    const quality = state.analysis.quality;
    $("qualityCheckGrid").innerHTML = quality.critical_checks
      .map(
        (check) => `<article class="quality-check ${check.status === "PASS" ? "" : "fail"}">
          <div><h3>${escapeHtml(check.label)}</h3><p>差异记录 ${formatNumber(
            check.difference_count,
            0,
          )}</p></div>
          <span class="quality-check-state">${escapeHtml(check.status)}</span>
        </article>`,
      )
      .join("");

    const anomalies = [
      [
        "未识别大区区域",
        state.analysis.meta.unknown_regions.length
          ? state.analysis.meta.unknown_regions.join("、")
          : "0",
      ],
      ["区间日期源行", quality.source_row_counts.date_range_source_rows],
      ["正需求但零发单单元", quality.positive_demand_no_issue.unit_count],
      ["未覆盖需求", `${formatNumber(quality.positive_demand_no_issue.demand_person_days)} 人日`],
      ["零需求残留发单", `${formatNumber(quality.zero_demand_with_issue.issued_person_days)} 人日`],
      ["超发单元", quality.over_issue.unit_count],
      ["累计超发", `${formatNumber(quality.over_issue.person_days)} 人日`],
      ["派遣超仓需单元", quality.over_dispatch.unit_count],
      ["测试单元剔除", quality.test_unit_count],
    ];
    $("anomalyList").innerHTML = anomalies
      .map(
        ([label, value]) =>
          `<div class="anomaly-item"><span>${label}</span><strong>${escapeHtml(value)}</strong></div>`,
      )
      .join("");

    const coverage = state.analysis.meta.coverage;
    const coverageRate = coverage.expected_days
      ? coverage.covered_days / coverage.expected_days
      : 0;
    $("coveragePanel").innerHTML = `
      <div class="coverage-meter" title="${formatRate(coverageRate)}"><span style="width:${
        coverageRate * 100
      }%"></span></div>
      <div class="coverage-meta"><span>已覆盖 ${coverage.covered_days} 天</span><strong>应覆盖 ${
        coverage.expected_days
      } 天</strong></div>
      ${
        coverage.complete
          ? '<div class="coverage-warning">统计期日期覆盖完整。重复区间仍按原文件事实叠加，不做静默去重。</div>'
          : `<div class="coverage-warning">缺少 ${coverage.missing_days.length} 天：${escapeHtml(
              coverage.missing_days.slice(0, 12).join("、"),
            )}${coverage.missing_days.length > 12 ? "…" : ""}</div>`
      }`;

    $("reconciliationBody").innerHTML = quality.month_region_reconciliation
      .map(
        (row) => `<tr>
          <td>${escapeHtml(row.month)}</td><td>${escapeHtml(row.region)}</td>
          <td>${formatNumber(row.basic_demand)}</td><td>${formatNumber(
            row.demand_detail_demand,
          )}</td>
          <td>${formatNumber(row.basic_issued)}</td><td>${formatNumber(
            row.demand_detail_issued,
          )}</td><td>${formatNumber(row.dispatch_detail_issued)}</td>
          <td>${formatNumber(row.basic_dispatched)}</td><td>${formatNumber(
            row.dispatch_detail_dispatched,
          )}</td>
          <td class="${row.status === "一致" ? "delta-positive" : "delta-negative"}">${escapeHtml(
            row.status,
          )}</td>
        </tr>`,
      )
      .join("");
  }

  function renderMethodView() {
    if (!state.analysis) return;
    const { method, meta } = state.analysis;
    const methodRows = [
      ["主指标", `${meta.primary_metric} = ${method.raw_rate}`],
      ["对照指标", `${meta.comparison_metric} = ${method.adjusted_rate}`],
      ["合理责任", method.reasonable_responsibility],
      ["仓库需求满足率", method.warehouse_rate],
      ["发单覆盖率", method.issue_coverage],
      ["统计期月均", method.average],
      ["统计期加权", method.period_weighted],
      ["原子粒度", method.atomic_grain],
      ["零需求残留发单", method.zero_demand_issue],
      ["测试数据", method.test_unit],
      ["无有效分母月份", meta.no_data_rule],
    ];
    $("methodContent").innerHTML = `
      <section class="method-section">
        <h3>指标职责边界</h3>
        <div class="method-callout">
          原始接单兑现率评价供应商对实际接单的兑现；超发校正率回答“扣除HR超发后，供应商对合理责任兑现多少”；
          仓库需求满足率只评价仓库人力保障，不能分摊后冒充单个供应商履约率。
        </div>
      </section>
      <section class="method-section">
        <h3>可审计计算规则</h3>
        <dl class="method-grid">
          ${methodRows
            .map(
              ([label, value]) =>
                `<dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd>`,
            )
            .join("")}
        </dl>
      </section>
      <section class="method-section">
        <h3>人日说明</h3>
        <div class="method-callout">
          1人工作1天记为1人日；同一个岗位每天需要5人、连续10天，就是50人日。需求、发单和派遣是事实人日；
          “合理责任”按多供应商发单比例折算，所以是责任人日等值，出现小数属于分摊结果，并不代表存在半个人。
        </div>
      </section>`;
  }

  function setupCanvas(canvas) {
    const rect = canvas.getBoundingClientRect();
    const width = Math.max(320, rect.width || canvas.parentElement.clientWidth || 640);
    const height = Math.max(240, rect.height || canvas.parentElement.clientHeight || 300);
    const ratio = Math.min(window.devicePixelRatio || 1, 2);
    canvas.width = Math.round(width * ratio);
    canvas.height = Math.round(height * ratio);
    const context = canvas.getContext("2d");
    context.setTransform(ratio, 0, 0, ratio, 0, 0);
    context.clearRect(0, 0, width, height);
    return { context, width, height };
  }

  function drawEmptyCanvas(context, width, height, message) {
    context.fillStyle = palette.muted;
    context.font = "12px Arial, sans-serif";
    context.textAlign = "center";
    context.fillText(message, width / 2, height / 2);
  }

  function drawTrendChart() {
    if (!state.analysis || $("tab-overview").classList.contains("is-active") === false) return;
    const canvas = $("trendChart");
    const { context, width, height } = setupCanvas(canvas);
    const display = contextDisplayRow();
    const warehouseMap = new Map(contextWarehouseRows().map((row) => [row.month, row]));
    if (!display) {
      drawEmptyCanvas(context, width, height, "当前筛选无趋势数据");
      return;
    }
    const series = [
      {
        color: palette.navy,
        values: display.by_month.map((row) => row.raw_rate),
      },
      {
        color: palette.orange,
        values: display.by_month.map((row) => row.adjusted_rate),
      },
      {
        color: palette.teal,
        values: display.by_month.map(
          (row) => warehouseMap.get(row.month)?.warehouse_fulfillment_rate ?? null,
        ),
      },
      {
        color: palette.gray,
        values: display.by_month.map(
          (row) => warehouseMap.get(row.month)?.issue_coverage_rate ?? null,
        ),
      },
    ];
    const values = series.flatMap((item) => item.values).filter((value) => value != null);
    if (!values.length) {
      drawEmptyCanvas(context, width, height, "当前筛选无趋势数据");
      return;
    }
    const margin = { top: 18, right: 18, bottom: 34, left: 46 };
    const plotWidth = width - margin.left - margin.right;
    const plotHeight = height - margin.top - margin.bottom;
    const maxValue = Math.max(1, Math.ceil(Math.max(...values) * 10) / 10);
    context.font = "10px Arial, sans-serif";
    context.lineWidth = 1;
    context.textAlign = "right";
    context.textBaseline = "middle";
    for (let tick = 0; tick <= 5; tick += 1) {
      const value = (maxValue * tick) / 5;
      const y = margin.top + plotHeight - (value / maxValue) * plotHeight;
      context.strokeStyle = palette.grid;
      context.beginPath();
      context.moveTo(margin.left, y);
      context.lineTo(width - margin.right, y);
      context.stroke();
      context.fillStyle = palette.muted;
      context.fillText(formatRate(value, 0), margin.left - 8, y);
    }
    const labels = state.analysis.meta.months.map((month) => month.label);
    const xAt = (index) =>
      labels.length === 1
        ? margin.left + plotWidth / 2
        : margin.left + (plotWidth * index) / (labels.length - 1);
    context.textAlign = "center";
    context.textBaseline = "top";
    labels.forEach((label, index) => {
      context.fillStyle = palette.ink;
      context.fillText(label, xAt(index), height - margin.bottom + 12);
    });
    series.forEach((item) => {
      context.strokeStyle = item.color;
      context.fillStyle = item.color;
      context.lineWidth = 2.2;
      context.beginPath();
      let started = false;
      item.values.forEach((value, index) => {
        if (value == null) {
          started = false;
          return;
        }
        const x = xAt(index);
        const y = margin.top + plotHeight - (value / maxValue) * plotHeight;
        if (!started) context.moveTo(x, y);
        else context.lineTo(x, y);
        started = true;
      });
      context.stroke();
      item.values.forEach((value, index) => {
        if (value == null) return;
        const x = xAt(index);
        const y = margin.top + plotHeight - (value / maxValue) * plotHeight;
        context.beginPath();
        context.arc(x, y, 3.6, 0, Math.PI * 2);
        context.fill();
        context.strokeStyle = "#ffffff";
        context.lineWidth = 1.5;
        context.stroke();
      });
    });
  }

  function scatterRows() {
    const { market, region, job } = selection();
    const targetType = job === "__all__" ? "supplier_total" : "supplier_job";
    return state.analysis.display_metrics
      .filter(
        (row) =>
          row.row_type === targetType &&
          (market === "__all__" || row.market === market) &&
          (region === "__all__" || row.region === region) &&
          (job === "__all__" || row.job === job) &&
          row.raw_average != null &&
          row.adjusted_average != null,
      )
      .map((row) => ({
        region: row.region,
        supplier: row.supplier_name,
        raw: row.raw_average,
        adjusted: row.adjusted_average,
        issued: row.totals.issued,
      }));
  }

  function drawSupplierScatter() {
    if (!state.analysis || $("tab-overview").classList.contains("is-active") === false) return;
    const canvas = $("supplierScatter");
    const { context, width, height } = setupCanvas(canvas);
    const rows = scatterRows();
    if (!rows.length) {
      drawEmptyCanvas(context, width, height, "当前筛选无供应商数据");
      return;
    }
    const margin = { top: 16, right: 16, bottom: 38, left: 46 };
    const plotWidth = width - margin.left - margin.right;
    const plotHeight = height - margin.top - margin.bottom;
    const maxValue = Math.max(
      1,
      Math.ceil(Math.max(...rows.flatMap((row) => [row.raw, row.adjusted])) * 10) / 10,
    );
    const xAt = (value) => margin.left + (Math.max(0, value) / maxValue) * plotWidth;
    const yAt = (value) =>
      margin.top + plotHeight - (Math.max(0, value) / maxValue) * plotHeight;
    context.font = "10px Arial, sans-serif";
    for (let tick = 0; tick <= 5; tick += 1) {
      const value = (maxValue * tick) / 5;
      const x = xAt(value);
      const y = yAt(value);
      context.strokeStyle = palette.grid;
      context.lineWidth = 1;
      context.beginPath();
      context.moveTo(x, margin.top);
      context.lineTo(x, height - margin.bottom);
      context.stroke();
      context.beginPath();
      context.moveTo(margin.left, y);
      context.lineTo(width - margin.right, y);
      context.stroke();
      context.fillStyle = palette.muted;
      context.textAlign = "center";
      context.textBaseline = "top";
      context.fillText(formatRate(value, 0), x, height - margin.bottom + 9);
      context.textAlign = "right";
      context.textBaseline = "middle";
      context.fillText(formatRate(value, 0), margin.left - 7, y);
    }
    context.strokeStyle = "#94a3b8";
    context.setLineDash([5, 4]);
    context.beginPath();
    context.moveTo(xAt(0), yAt(0));
    context.lineTo(xAt(maxValue), yAt(maxValue));
    context.stroke();
    context.setLineDash([]);
    const regionColor = new Map(
      state.analysis.meta.regions.map((region, index) => [
        region,
        palette.region[index % palette.region.length],
      ]),
    );
    const maxIssued = Math.max(...rows.map((row) => row.issued), 1);
    rows.forEach((row) => {
      const radius = 3.5 + 8 * Math.sqrt(row.issued / maxIssued);
      context.globalAlpha = 0.72;
      context.fillStyle = regionColor.get(row.region) || palette.navy;
      context.beginPath();
      context.arc(xAt(row.raw), yAt(row.adjusted), radius, 0, Math.PI * 2);
      context.fill();
      context.globalAlpha = 1;
      context.strokeStyle = "#ffffff";
      context.lineWidth = 1;
      context.stroke();
    });
    context.fillStyle = palette.ink;
    context.font = "10px Arial, sans-serif";
    context.textAlign = "center";
    context.fillText("原始月均", margin.left + plotWidth / 2, height - 3);
    context.save();
    context.translate(11, margin.top + plotHeight / 2);
    context.rotate(-Math.PI / 2);
    context.fillText("校正月均", 0, 0);
    context.restore();
  }

  function drawOverviewCharts() {
    drawTrendChart();
    drawSupplierScatter();
  }

  function activateTab(tabName) {
    state.activeTab = tabName;
    document.querySelectorAll(".tab").forEach((button) => {
      button.classList.toggle("is-active", button.dataset.tab === tabName);
    });
    document.querySelectorAll(".tab-panel").forEach((panel) => {
      panel.classList.toggle("is-active", panel.id === `tab-${tabName}`);
    });
    if (tabName === "overview") {
      window.requestAnimationFrame(drawOverviewCharts);
    }
  }

  async function exportReport() {
    if (!state.analysis) return;
    if (triggerReadyDownload()) {
      showToast(`正在下载 ${state.downloadFileName}`);
      return;
    }
    setBusy(true, "正在生成含公式链与十张工作表的严谨版报表…");
    try {
      const result = await exporter.buildReport(state.analysis);
      prepareDownload(result.blob, result.fileName);
      triggerReadyDownload();
      setStatus(`报表已生成；若未自动下载，请点击“再次点击下载”`, "success");
      showToast(`报表已生成，可再次点击下载`);
    } catch (error) {
      console.error(error);
      setStatus(error.message || "报表导出失败", "error");
      showToast(error.message || "报表导出失败", 7000);
    } finally {
      setBusy(false);
    }
  }

  function handleFilterChange() {
    if (!state.analysis) return;
    renderOverview();
    renderSupplierTable();
    renderWarehouseView();
    window.requestAnimationFrame(drawOverviewCharts);
  }

  function handleMarketChange() {
    if (!state.analysis) return;
    populateRegionFilter();
    handleFilterChange();
  }

  elements.fileInput.addEventListener("change", (event) => {
    addFiles(event.target.files);
    event.target.value = "";
  });
  elements.dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    elements.dropzone.classList.add("is-dragging");
  });
  elements.dropzone.addEventListener("dragleave", () => {
    elements.dropzone.classList.remove("is-dragging");
  });
  elements.dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    elements.dropzone.classList.remove("is-dragging");
    addFiles(event.dataTransfer.files);
  });
  elements.fileList.addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-file]");
    if (!button) return;
    state.files.splice(Number(button.dataset.removeFile), 1);
    state.analysis = null;
    renderFileList();
    resetAnalysisPresentation();
  });
  elements.clearFilesButton.addEventListener("click", clearFiles);
  elements.analyzeButton.addEventListener("click", analyze);
  elements.exportButton.addEventListener("click", exportReport);
  elements.marketFilter.addEventListener("change", handleMarketChange);
  elements.regionFilter.addEventListener("change", handleFilterChange);
  elements.jobFilter.addEventListener("change", handleFilterChange);
  elements.supplierSearch.addEventListener("input", () => {
    renderSupplierPreview();
    renderSupplierTable();
  });
  document.querySelector(".tabs").addEventListener("click", (event) => {
    const button = event.target.closest("[data-tab]");
    if (button) activateTab(button.dataset.tab);
  });
  document.querySelector(".segmented-control").addEventListener("click", (event) => {
    const button = event.target.closest("[data-supplier-metric]");
    if (!button) return;
    state.supplierMetric = button.dataset.supplierMetric;
    document.querySelectorAll("[data-supplier-metric]").forEach((item) => {
      item.classList.toggle("is-active", item === button);
    });
    renderSupplierTable();
  });
  window.addEventListener("resize", () => {
    window.clearTimeout(state.resizeTimer);
    state.resizeTimer = window.setTimeout(drawOverviewCharts, 120);
  });
  window.addEventListener("beforeunload", clearReadyDownload);

  renderFileList();
})();

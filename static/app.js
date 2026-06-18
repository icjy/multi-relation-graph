const state = {
  datasetId: "",
  centerType: "borrower",
  centerId: "",
  overdueBasis: "any",
  communityMethod: "connected_components",
  reloanFilter: "all",
  returnFilter: "all",
  lastFile: null,
  lastFileName: "",
  currentGraph: null,
  featureTable: { columns: [], rows: [], page: 1, pageSize: 50, communityFilter: "" },
  complexQuery: { columns: [], rows: [], summary: {}, input_counts: {}, matched_value_counts: {}, unmatched_values: {}, page: 1, pageSize: 20 },
  cluster: {
    mode: "full",
    pageSize: 50,
    full: {
      jobId: "",
      result: null,
      detailPage: 1,
      summaryPage: 1,
      detailSort: { column: "addr_cluster_id", direction: "asc" },
      summarySort: { column: "count", direction: "desc" },
      clusterFilter: "",
    },
    incremental: {
      jobId: "",
      result: null,
      detailPage: 1,
      summaryPage: 1,
      detailSort: { column: "addr_cluster_id", direction: "asc" },
      summarySort: { column: "count", direction: "desc" },
      clusterFilter: "",
    },
  },
};
const graphHiddenTypes = new Set();

const fileInput = document.querySelector("#fileInput");
const dropZone = document.querySelector("#dropZone");
const overdueBasis = document.querySelector("#overdueBasis");
const communityMethod = document.querySelector("#communityMethod");
const reloanFilter = document.querySelector("#reloanFilter");
const returnFilter = document.querySelector("#returnFilter");
const borrowerMode = document.querySelector("#borrowerMode");
const agentMode = document.querySelector("#agentMode");
const centerInput = document.querySelector("#centerInput");
const queryButton = document.querySelector("#queryButton");
const exportCsvButton = document.querySelector("#exportCsvButton");
const exportExcelButton = document.querySelector("#exportExcelButton");
const featureCommunityFilter = document.querySelector("#featureCommunityFilter");
const analysisProgress = document.querySelector("#analysisProgress");
const progressTitle = document.querySelector("#progressTitle");
const progressText = document.querySelector("#progressText");
const progressBar = document.querySelector("#progressBar");
const navButtons = document.querySelectorAll(".nav-button");
const views = document.querySelectorAll(".view");
const appShell = document.querySelector(".app-shell");
const navToggle = document.querySelector("#navToggle");
const navExpand = document.querySelector("#navExpand");
const queryReloanFilter = document.querySelector("#queryReloanFilter");
const queryReturnFilter = document.querySelector("#queryReturnFilter");
const queryFinalResultFilter = document.querySelector("#queryFinalResultFilter");
const queryBorrowerValues = document.querySelector("#queryBorrowerValues");
const queryAgentValues = document.querySelector("#queryAgentValues");
const queryDeviceValues = document.querySelector("#queryDeviceValues");
const queryIpValues = document.querySelector("#queryIpValues");
const queryAddrValues = document.querySelector("#queryAddrValues");
const runComplexQueryButton = document.querySelector("#runComplexQuery");
const clearComplexQueryButton = document.querySelector("#clearComplexQuery");
const queryPageSize = document.querySelector("#queryPageSize");
const fullClusterFile = document.querySelector("#fullClusterFile");
const baseClusterFile = document.querySelector("#baseClusterFile");
const incrementalClusterFile = document.querySelector("#incrementalClusterFile");
const fullClusterFileStatus = document.querySelector("#fullClusterFileStatus");
const baseClusterFileStatus = document.querySelector("#baseClusterFileStatus");
const incrementalClusterFileStatus = document.querySelector("#incrementalClusterFileStatus");
const runFullClusterButton = document.querySelector("#runFullCluster");
const runIncrementalClusterButton = document.querySelector("#runIncrementalCluster");
const fullClusterTab = document.querySelector("#fullClusterTab");
const incrementalClusterTab = document.querySelector("#incrementalClusterTab");
const fullClusterPane = document.querySelector("#fullClusterPane");
const incrementalClusterPane = document.querySelector("#incrementalClusterPane");
const clusterProgress = document.querySelector("#clusterProgress");
const clusterProgressTitle = document.querySelector("#clusterProgressTitle");
const clusterProgressText = document.querySelector("#clusterProgressText");
const clusterProgressBar = document.querySelector("#clusterProgressBar");
const clusterDetailFilter = document.querySelector("#clusterDetailFilter");
const clusterIndexType = document.querySelector("#clusterIndexType");
const exportClusterDetailCsv = document.querySelector("#exportClusterDetailCsv");
const exportClusterDetailExcel = document.querySelector("#exportClusterDetailExcel");
const exportClusterSummaryCsv = document.querySelector("#exportClusterSummaryCsv");
const exportClusterSummaryExcel = document.querySelector("#exportClusterSummaryExcel");
const exportClusterEngineering = document.querySelector("#exportClusterEngineering");
let progressHideTimer = 0;
let progressPulseTimer = 0;
let currentProgress = 0;
let clusterPollTimer = 0;

navButtons.forEach((button) => {
  button.addEventListener("click", () => switchView(button.dataset.view));
});
navToggle?.addEventListener("click", () => setNavHidden(true));
navExpand?.addEventListener("click", () => setNavHidden(false));

fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) analyzeFile(fileInput.files[0]);
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragging");
});

dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragging"));

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragging");
  if (event.dataTransfer.files[0]) analyzeFile(event.dataTransfer.files[0]);
});

overdueBasis.addEventListener("change", () => {
  state.overdueBasis = overdueBasis.value;
  if (state.lastFile) analyzeFile(state.lastFile);
});

communityMethod?.addEventListener("change", () => {
  state.communityMethod = communityMethod.value;
  if (state.lastFile) analyzeFile(state.lastFile);
});

reloanFilter?.addEventListener("change", () => {
  state.reloanFilter = reloanFilter.value;
  if (state.lastFile) analyzeFile(state.lastFile);
});

returnFilter?.addEventListener("change", () => {
  state.returnFilter = returnFilter.value;
  if (state.lastFile) analyzeFile(state.lastFile);
});

borrowerMode.addEventListener("click", () => setCenterType("borrower"));
agentMode.addEventListener("click", () => setCenterType("agent"));

queryButton.addEventListener("click", queryGraph);
centerInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") queryGraph();
});

centerInput.addEventListener("input", debounce(searchCenters, 240));
exportCsvButton?.addEventListener("click", () => exportFeatureTable("csv"));
exportExcelButton?.addEventListener("click", () => exportFeatureTable("xls"));
featureCommunityFilter?.addEventListener("input", debounce(() => {
  state.featureTable.communityFilter = featureCommunityFilter.value.trim();
  state.featureTable.page = 1;
  renderFeatureTable(state.featureTable);
}, 180));
runComplexQueryButton?.addEventListener("click", runComplexQuery);
clearComplexQueryButton?.addEventListener("click", clearComplexQuery);
fullClusterFile?.addEventListener("change", () => updateFileStatus(fullClusterFile, fullClusterFileStatus, "未选择文件"));
baseClusterFile?.addEventListener("change", () => updateFileStatus(baseClusterFile, baseClusterFileStatus, "未选择工程文件"));
incrementalClusterFile?.addEventListener("change", () => updateFileStatus(incrementalClusterFile, incrementalClusterFileStatus, "未选择增量数据"));
runFullClusterButton?.addEventListener("click", runFullCluster);
runIncrementalClusterButton?.addEventListener("click", runIncrementalCluster);
fullClusterTab?.addEventListener("click", () => setClusterMode("full"));
incrementalClusterTab?.addEventListener("click", () => setClusterMode("incremental"));
clusterIndexType?.addEventListener("change", updateIncrementalParamVisibility);
clusterDetailFilter?.addEventListener("input", debounce(() => {
  const scope = currentClusterState();
  scope.clusterFilter = clusterDetailFilter.value.trim();
  scope.detailPage = 1;
  renderClusterDetailTable();
}, 180));
exportClusterDetailCsv?.addEventListener("click", () => exportClusterTable("detail", "csv"));
exportClusterDetailExcel?.addEventListener("click", () => exportClusterTable("detail", "xls"));
exportClusterSummaryCsv?.addEventListener("click", () => exportClusterTable("summary", "csv"));
exportClusterSummaryExcel?.addEventListener("click", () => exportClusterTable("summary", "xls"));
exportClusterEngineering?.addEventListener("click", exportClusterEngineeringFile);
queryPageSize?.addEventListener("change", () => {
  state.complexQuery.pageSize = Number(queryPageSize.value) || 20;
  state.complexQuery.page = 1;
  renderComplexQueryResult();
});

function setCenterType(type) {
  state.centerType = type;
  borrowerMode.classList.toggle("active", type === "borrower");
  agentMode.classList.toggle("active", type === "agent");
}

function setClusterMode(mode) {
  const isFull = mode === "full";
  state.cluster.mode = isFull ? "full" : "incremental";
  fullClusterTab?.classList.toggle("active", isFull);
  incrementalClusterTab?.classList.toggle("active", !isFull);
  if (fullClusterPane) fullClusterPane.hidden = !isFull;
  if (incrementalClusterPane) incrementalClusterPane.hidden = isFull;
  const scope = currentClusterState();
  if (clusterDetailFilter) clusterDetailFilter.value = scope.clusterFilter || "";
  renderClusterResult();
  updateIncrementalParamVisibility();
}

function currentClusterState() {
  return state.cluster[state.cluster.mode] || state.cluster.full;
}

function updateFileStatus(input, target, emptyText) {
  if (!target) return;
  const file = input?.files?.[0];
  if (!file) {
    target.textContent = emptyText;
    target.classList.remove("loaded");
    return;
  }
  target.textContent = `已选择：${file.name}（${formatFileSize(file.size)}）`;
  target.classList.add("loaded");
}

function updateIncrementalParamVisibility() {
  const type = String(clusterIndexType?.value || "FlatL2").toLowerCase();
  document.querySelectorAll(".index-param-hnsw").forEach((item) => {
    item.hidden = type !== "hnsw";
  });
  document.querySelectorAll(".index-param-ivf").forEach((item) => {
    item.hidden = type !== "ivf";
  });
}

function switchView(viewId) {
  navButtons.forEach((button) => button.classList.toggle("active", button.dataset.view === viewId));
  views.forEach((view) => {
    const isActive = view.id === viewId;
    view.classList.toggle("active", isActive);
    view.hidden = !isActive;
  });
  if (viewId === "addressCluster") renderClusterResult();
}

switchView("graphView");
updateIncrementalParamVisibility();

function setNavHidden(hidden) {
  appShell?.classList.toggle("nav-hidden", hidden);
  if (navExpand) navExpand.hidden = !hidden;
}

async function analyzeFile(file) {
  if (state.lastFileName && state.lastFileName !== file.name) {
    state.reloanFilter = "all";
    if (reloanFilter) reloanFilter.value = "all";
  }
  state.lastFile = file;
  state.lastFileName = file.name;
  showProgress(6, "准备分析", "正在读取文件");
  showToast("正在解析文件并计算二度指标...");
  try {
    const contentBase64 = await fileToBase64(file);
    showProgress(28, "上传数据", "正在提交筛选条件和社区算法");
    startProgressPulse(76, "计算中", "后端正在构建图谱、计算社区和特征宽表");
    const response = await postJson("/api/analyze", {
      filename: file.name,
      content_base64: contentBase64,
      overdue_basis: state.overdueBasis,
      community_method: state.communityMethod,
      reloan_filter: state.reloanFilter,
      return_filter: state.returnFilter,
    });
    stopProgressPulse();
    showProgress(82, "渲染结果", "正在刷新图谱、社区表和特征宽表");
    state.datasetId = response.dataset_id;
    state.centerType = response.graph.center?.type || "borrower";
    state.centerId = response.graph.center?.id || "";
    setCenterType(state.centerType);
    centerInput.value = state.centerId;
    renderAll(response);
    completeProgress("分析完成");
    showToast("分析完成");
  } catch (error) {
    stopProgressPulse();
    hideProgress();
    showToast(error.message || String(error));
  }
}

async function queryGraph() {
  if (!state.datasetId) {
    showToast("请先上传数据文件");
    return;
  }
  const centerId = centerInput.value.trim();
  if (!centerId) {
    showToast("请输入中心节点");
    return;
  }
  try {
    const response = await postJson("/api/graph", {
      dataset_id: state.datasetId,
      center_type: state.centerType,
      center_id: centerId,
      overdue_basis: state.overdueBasis,
    });
    state.centerId = response.graph.center.id;
    renderGraph(response.graph);
    renderMetrics(response.graph);
  } catch (error) {
    showToast(error.message || String(error));
  }
}

async function searchCenters() {
  if (!state.datasetId) return;
  const query = centerInput.value.trim();
  const container = document.querySelector("#searchResults");
  if (!query) {
    container.innerHTML = "";
    return;
  }
  try {
    const response = await postJson("/api/search", {
      dataset_id: state.datasetId,
      query,
    });
    const hits = [...response.borrowers, ...response.agents].slice(0, 8);
    container.innerHTML = hits.map((hit) => `
      <button class="search-hit" type="button" data-type="${hit.type}" data-id="${hit.id}">
        ${hit.type === "agent" ? "中介" : "借款人"} ${hit.label}
      </button>
    `).join("");
    container.querySelectorAll("button").forEach((button) => {
      button.addEventListener("click", () => {
        setCenterType(button.dataset.type);
        centerInput.value = button.dataset.id;
        queryGraph();
      });
    });
  } catch {
    container.innerHTML = "";
  }
}

function renderAll(data) {
  const currentFeatureState = state.featureTable || {};
  state.featureTable = {
    ...(data.feature_table || { columns: [], rows: [] }),
    page: 1,
    pageSize: currentFeatureState.pageSize || 50,
    communityFilter: currentFeatureState.communityFilter || "",
  };
  if (featureCommunityFilter) featureCommunityFilter.value = state.featureTable.communityFilter;
  updateFilterOptions(data.filter_options || {});
  renderSummary(data.summary);
  renderGraph(data.graph);
  renderMetrics(data.graph);
  renderCommunityTable(data.communities);
  renderAgentTable(data.top_agents);
  renderBorrowerTable(data.top_borrowers);
  renderFeatureTable(state.featureTable);
}

function updateFilterOptions(options) {
  updateReLoanOptions(options.reloan || []);
  updateSelectOptions(queryReturnFilter, options.return || [], "all", returnLabel);
  updateSelectOptions(queryFinalResultFilter, options.final_result || [], "all", finalResultLabel);
}

function updateReLoanOptions(values) {
  if (!reloanFilter) return;
  const selected = state.reloanFilter || "all";
  updateSelectOptions(reloanFilter, values, selected, reloanLabel);
  state.reloanFilter = values.map(String).includes(selected) ? selected : "all";
  reloanFilter.value = state.reloanFilter;
  updateSelectOptions(queryReloanFilter, values, "all", reloanLabel);
}

function updateSelectOptions(select, values, selected = "all", labeler = (value) => value) {
  if (!select) return;
  const options = [
    ["all", "全部"],
    ...values.map((value) => [String(value), labeler(value)]),
  ];
  select.innerHTML = options
    .map(([value, label]) => `<option value="${escapeHtml(value)}">${escapeHtml(label)}</option>`)
    .join("");
  select.value = values.map(String).includes(selected) ? selected : "all";
}

function reloanLabel(value) {
  const text = String(value ?? "");
  if (text === "0") return "首贷 / 0";
  if (text === "1") return "复贷 / 1";
  return text || "空值";
}

function returnLabel(value) {
  const text = String(value ?? "");
  if (text === "0") return "未退货 / 0";
  if (text === "1") return "已退货 / 1";
  return text || "空值";
}

function finalResultLabel(value) {
  const text = String(value ?? "");
  if (text === "30") return "借款通过 / 30";
  return text ? `未通过 / ${text}` : "空值";
}

function showProgress(percent, title, detail) {
  if (!analysisProgress) return;
  window.clearTimeout(progressHideTimer);
  analysisProgress.hidden = false;
  const value = Math.max(0, Math.min(100, Number(percent) || 0));
  currentProgress = value;
  progressTitle.textContent = title;
  progressText.textContent = detail;
  progressBar.style.width = `${value}%`;
  analysisProgress.querySelector(".progress-track")?.setAttribute("aria-valuenow", String(Math.round(value)));
}

function startProgressPulse(target, title, detail) {
  stopProgressPulse();
  progressPulseTimer = window.setInterval(() => {
    if (currentProgress >= target) return;
    const next = currentProgress + Math.max(0.6, (target - currentProgress) * 0.08);
    showProgress(Math.min(target, next), title, detail);
  }, 420);
}

function stopProgressPulse() {
  window.clearInterval(progressPulseTimer);
}

function completeProgress(detail = "已完成") {
  stopProgressPulse();
  showProgress(100, "分析完成", detail);
  progressHideTimer = window.setTimeout(hideProgress, 700);
}

function hideProgress() {
  if (!analysisProgress) return;
  window.clearTimeout(progressHideTimer);
  stopProgressPulse();
  progressBar.style.width = "0%";
  analysisProgress.hidden = true;
}

function renderSummary(summary) {
  const methodLabel = {
    connected_components: "连通分量",
    louvain: "Louvain",
    leiden: "Leiden",
  }[summary.community_method || state.communityMethod] || "连通分量";
  const items = [
    ["上传原始行数", formatNumber(summary.uploaded_row_count ?? summary.row_count)],
    ["参与分析行数", formatNumber(summary.analyzed_row_count ?? summary.row_count)],
    ["有效贷款数", formatNumber(summary.valid_loan_count)],
    ["中介数", formatNumber(summary.agent_count)],
    ["借款人数", formatNumber(summary.borrower_count)],
    ["关系数", formatNumber(summary.relation_count)],
    ["社区算法", methodLabel],
    ["疑似团伙数", formatNumber(summary.community_count)],
    [summary.community_size_label || "最大团伙规模", formatNumber(summary.max_community_size ?? summary.max_community_agents)],
    ["整体逾期率", formatPercent(summary.overdue_rate)],
  ];
  if (summary.skipped_rows) items.push(["跳过行数", formatNumber(summary.skipped_rows), "warning"]);
  if (summary.missing_columns?.length) items.push(["缺失字段", summary.missing_columns.length, "warning"]);
  document.querySelector("#summaryCards").innerHTML = items.map(([label, value, tone]) => `
    <div class="summary-card ${tone || ""}">
      <span>${label}</span>
      <strong>${value}</strong>
    </div>
  `).join("");
}

function renderMetrics(graph) {
  const metrics = graph.metrics || {};
  const isAgent = graph.center?.type === "agent";
  const items = isAgent
    ? [
        ["中介手机号", mask(graph.center.id)],
        ["一度借款人数", formatNumber(metrics.first_degree_borrowers)],
        ["二度中介数", formatNumber(metrics.second_degree_agents)],
      ]
    : [
        ["借款人", mask(graph.center?.id || "")],
        ["一度中介数", formatNumber(metrics.first_degree_agents)],
        ["本人贷款数", formatNumber(metrics.loan_count)],
        ["本人逾期笔数", formatNumber(metrics.own_overdue_loans)],
        ["二度借款人数", formatNumber(metrics.second_degree_borrowers)],
        ["二度贷款数", formatNumber(metrics.second_degree_loans)],
        ["二度逾期笔数", formatNumber(metrics.second_degree_overdue_loans)],
        ["二度借款逾期率", formatPercent(metrics.second_degree_overdue_rate)],
      ];

  document.querySelector("#centerMetrics").innerHTML = items.map(([label, value]) => `
    <div class="metric-item">
      <span>${label}</span>
      <strong>${value}</strong>
    </div>
  `).join("");
}

function renderFeatureTable(featureTable) {
  const container = document.querySelector("#featureTable");
  const pagination = document.querySelector("#featurePagination");
  const subtitle = document.querySelector("#featureSubtitle");
  if (!container || !pagination) return;
  const columns = featureTable?.columns || [];
  const rows = filteredFeatureRows(featureTable);
  const rawRows = featureTable?.rows || [];
  const pageSize = Number(featureTable.pageSize || 50);
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  const page = Math.min(Math.max(1, Number(featureTable.page || 1)), totalPages);
  featureTable.page = page;
  const filterText = String(featureTable.communityFilter || "").trim();
  subtitle.textContent = rawRows.length
    ? `共 ${formatNumber(rawRows.length)} 行，${formatNumber(columns.length)} 列；${filterText ? `筛选命中 ${formatNumber(rows.length)} 行；` : ""}当前第 ${formatNumber(page)} / ${formatNumber(totalPages)} 页。`
    : "上传数据后生成原始字段 + 加工特征";
  if (!columns.length || !rawRows.length) {
    container.innerHTML = "暂无数据";
    container.classList.add("empty");
    pagination.innerHTML = "";
    return;
  }
  if (!rows.length) {
    container.innerHTML = "无匹配结果";
    container.classList.add("empty");
    pagination.innerHTML = "";
    return;
  }
  container.classList.remove("empty");
  const start = (page - 1) * pageSize;
  const pageRows = rows.slice(start, start + pageSize);
  container.innerHTML = `
    <table>
      <thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr></thead>
      <tbody>
        ${pageRows.map((row) => `
          <tr>${columns.map((column) => `<td>${formatCell(row[column])}</td>`).join("")}</tr>
        `).join("")}
      </tbody>
    </table>
  `;
  pagination.innerHTML = `
    <button type="button" data-page="prev" ${page <= 1 ? "disabled" : ""}>上一页</button>
    <span>第 ${formatNumber(page)} / ${formatNumber(totalPages)} 页</span>
    <label class="page-jump">跳至 <input type="number" min="1" max="${totalPages}" value="${page}" aria-label="跳转页码" /> 页</label>
    <button type="button" data-page="next" ${page >= totalPages ? "disabled" : ""}>下一页</button>
  `;
  pagination.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", () => {
      featureTable.page += button.dataset.page === "prev" ? -1 : 1;
      renderFeatureTable(featureTable);
    });
  });
  bindPaginationJump(pagination, totalPages, (targetPage) => {
    featureTable.page = targetPage;
    renderFeatureTable(featureTable);
  });
}

function filteredFeatureRows(featureTable) {
  const rows = featureTable?.rows || [];
  const filterText = String(featureTable?.communityFilter || "").trim();
  if (!filterText) return rows;
  const columns = featureTable?.columns || [];
  const communityColumn = columns.includes("社区id(Agent)")
    ? "社区id(Agent)"
    : columns.find((column) => String(column).toLowerCase() === "社区id(agent)")
      || columns.find((column) => String(column).startsWith("社区id"));
  if (!communityColumn) return [];
  return rows.filter((row) => String(row[communityColumn] ?? "") === filterText);
}

async function runComplexQuery() {
  if (!state.datasetId) {
    showToast("请先在图谱分析页上传数据文件");
    return;
  }
  try {
    const response = await postJson("/api/complex-query", {
      dataset_id: state.datasetId,
      filters: {
        reloan_flag: queryReloanFilter?.value || "all",
        return_flag: queryReturnFilter?.value || "all",
        final_result: queryFinalResultFilter?.value || "all",
      },
      query_values: getComplexQueryValues(),
    });
    state.complexQuery = {
      ...response,
      page: 1,
      pageSize: Number(queryPageSize?.value) || 20,
    };
    renderComplexQueryResult();
    showToast(`查询完成，命中 ${formatNumber(response.summary?.row_count || 0)} 行`);
  } catch (error) {
    showToast(error.message || String(error));
  }
}

function getComplexQueryValues() {
  return {
    app_user_id: queryBorrowerValues?.value || "",
    consigneeMobileId: queryAgentValues?.value || "",
    device_id: queryDeviceValues?.value || "",
    ip: queryIpValues?.value || "",
    addr_cluster_id: queryAddrValues?.value || "",
  };
}

function clearComplexQuery() {
  [queryBorrowerValues, queryAgentValues, queryDeviceValues, queryIpValues, queryAddrValues].forEach((input) => {
    if (input) input.value = "";
  });
  if (queryReloanFilter) queryReloanFilter.value = "all";
  if (queryReturnFilter) queryReturnFilter.value = "all";
  if (queryFinalResultFilter) queryFinalResultFilter.value = "all";
  state.complexQuery = { columns: [], rows: [], summary: {}, input_counts: {}, matched_value_counts: {}, unmatched_values: {}, page: 1, pageSize: Number(queryPageSize?.value) || 20 };
  renderComplexQueryResult();
}

function renderComplexQueryResult() {
  const result = state.complexQuery || {};
  renderQuerySummary(result.summary || {});
  renderQueryHitDetails(result);
  renderQueryResultTable(result);
}

function renderQuerySummary(summary) {
  const container = document.querySelector("#querySummary");
  if (!container) return;
  const items = [
    ["命中行数", summary.row_count],
    ["命中 loan_task_id 数", summary.loan_task_id_count],
    ["用户数 app_user_id", summary.borrower_count],
    ["收货人手机号数", summary.agent_count],
    ["设备数 device_id", summary.device_count],
    ["IP数 ip", summary.ip_count],
    ["地址簇数 addr_cluster_id", summary.addr_cluster_count],
    ["原始地址数 consigneeAddr", summary.raw_address_count],
    ["预处理地址数 receiverAddr", summary.clean_address_count],
  ];
  container.innerHTML = items.map(([label, value]) => `
    <div class="summary-card">
      <span>${label}</span>
      <strong>${formatNumber(value || 0)}</strong>
    </div>
  `).join("");
}

function renderQueryHitDetails(result) {
  const container = document.querySelector("#queryHitDetails");
  const subtitle = document.querySelector("#queryHitSubtitle");
  if (!container) return;
  const labels = {
    app_user_id: "用户 app_user_id",
    consigneeMobileId: "收货手机号 consigneeMobileId",
    device_id: "设备 device_id",
    ip: "IP",
    addr_cluster_id: "地址簇 addr_cluster_id",
  };
  const fields = Object.keys(labels);
  const inputCounts = result.input_counts || {};
  const matchedCounts = result.matched_value_counts || {};
  const unmatched = result.unmatched_values || {};
  const hasInput = fields.some((field) => inputCounts[field]);
  subtitle.textContent = hasInput
    ? `命中 ${formatNumber(result.summary?.row_count || 0)} 行；以下按输入字段展示命中值和未命中值。`
    : "未输入实体值时，仅按下拉条件返回命中行。";
  if (!hasInput) {
    container.innerHTML = "暂无输入值统计";
    container.classList.add("empty");
    return;
  }
  container.classList.remove("empty");
  container.innerHTML = fields
    .filter((field) => inputCounts[field])
    .map((field) => {
      const missing = unmatched[field] || [];
      return `
        <div class="query-hit-row">
          <strong>${labels[field]}</strong>
          <span>
            输入 ${formatNumber(inputCounts[field])} 个，命中 ${formatNumber(matchedCounts[field] || 0)} 个，未命中 ${formatNumber(missing.length)} 个
            ${missing.length ? `<div class="miss-list">${missing.map((item) => `<code>${escapeHtml(item)}</code>`).join("")}</div>` : ""}
          </span>
        </div>
      `;
    }).join("");
}

function renderQueryResultTable(result) {
  const container = document.querySelector("#queryResultTable");
  const pagination = document.querySelector("#queryPagination");
  const subtitle = document.querySelector("#queryTableSubtitle");
  if (!container || !pagination) return;
  const columns = result.columns || [];
  const rows = result.rows || [];
  const pageSize = Number(result.pageSize || 20);
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  const page = Math.min(Math.max(1, Number(result.page || 1)), totalPages);
  state.complexQuery.page = page;
  subtitle.textContent = rows.length
    ? `共 ${formatNumber(rows.length)} 行，${formatNumber(columns.length)} 列；当前第 ${formatNumber(page)} / ${formatNumber(totalPages)} 页。`
    : "展示上传文件原始列对应的命中行。";
  if (!columns.length || !rows.length) {
    container.innerHTML = "暂无查询";
    container.classList.add("empty");
    pagination.innerHTML = "";
    return;
  }
  container.classList.remove("empty");
  const start = (page - 1) * pageSize;
  const pageRows = rows.slice(start, start + pageSize);
  container.innerHTML = `
    <table>
      <thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr></thead>
      <tbody>
        ${pageRows.map((row) => `
          <tr>${columns.map((column) => `<td>${formatCell(row[column])}</td>`).join("")}</tr>
        `).join("")}
      </tbody>
    </table>
  `;
  pagination.innerHTML = `
    <button type="button" data-page="prev" ${page <= 1 ? "disabled" : ""}>上一页</button>
    <span>第 ${formatNumber(page)} / ${formatNumber(totalPages)} 页</span>
    <label class="page-jump">跳至 <input type="number" min="1" max="${totalPages}" value="${page}" aria-label="跳转页码" /> 页</label>
    <button type="button" data-page="next" ${page >= totalPages ? "disabled" : ""}>下一页</button>
  `;
  pagination.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", () => {
      state.complexQuery.page += button.dataset.page === "prev" ? -1 : 1;
      renderQueryResultTable(state.complexQuery);
    });
  });
  bindPaginationJump(pagination, totalPages, (targetPage) => {
    state.complexQuery.page = targetPage;
    renderQueryResultTable(state.complexQuery);
  });
}

async function runFullCluster() {
  const file = fullClusterFile?.files?.[0];
  if (!file) {
    showToast("请先选择全量聚类数据");
    return;
  }
  try {
    showClusterProgress(5, "准备全量聚类", "正在读取文件");
    const contentBase64 = await fileToBase64(file);
    const response = await postJson("/api/cluster/full", {
      filename: file.name,
      content_base64: contentBase64,
      params: fullClusterParams(),
    });
    state.cluster.full.jobId = response.job_id;
    pollClusterJob(response.job_id, "full");
  } catch (error) {
    hideClusterProgress();
    showToast(error.message || String(error));
  }
}

async function runIncrementalCluster() {
  const baseFile = baseClusterFile?.files?.[0];
  const incrementalFile = incrementalClusterFile?.files?.[0];
  if (!baseFile || !incrementalFile) {
    showToast("请同时选择工程文件和增量数据");
    return;
  }
  try {
    showClusterProgress(5, "准备增量聚类", "正在读取工程文件和增量数据");
    const [baseContent, incrementalContent] = await Promise.all([
      fileToBase64(baseFile),
      fileToBase64(incrementalFile),
    ]);
    const response = await postJson("/api/cluster/incremental", {
      base_filename: baseFile.name,
      base_content_base64: baseContent,
      incremental_filename: incrementalFile.name,
      incremental_content_base64: incrementalContent,
      params: incrementalClusterParams(),
    });
    state.cluster.incremental.jobId = response.job_id;
    pollClusterJob(response.job_id, "incremental");
  } catch (error) {
    hideClusterProgress();
    showToast(error.message || String(error));
  }
}

function fullClusterParams() {
  return {
    model: document.querySelector("#clusterModel")?.value || "paraphrase-multilingual-MiniLM-L12-v2",
    eps: Number(document.querySelector("#clusterEps")?.value || 0.68),
    min_samples: Number(document.querySelector("#clusterMinSamples")?.value || 1),
    metric: document.querySelector("#clusterMetric")?.value || "euclidean",
    normalize: Boolean(document.querySelector("#clusterNormalize")?.checked),
  };
}

function incrementalClusterParams() {
  return {
    index_type: document.querySelector("#clusterIndexType")?.value || "FlatL2",
    threshold: Number(document.querySelector("#clusterThreshold")?.value || 0.68),
    normalize: Boolean(document.querySelector("#clusterNormalize")?.checked),
    hnsw_m: Number(document.querySelector("#clusterHnswM")?.value || 32),
    hnsw_ef_search: Number(document.querySelector("#clusterHnswEf")?.value || 64),
    ivf_nlist: Number(document.querySelector("#clusterIvfNlist")?.value || 64),
    ivf_nprobe: Number(document.querySelector("#clusterIvfNprobe")?.value || 8),
  };
}

async function pollClusterJob(jobId, mode) {
  window.clearTimeout(clusterPollTimer);
  try {
    const status = await postJson("/api/cluster/status", { job_id: jobId });
    showClusterProgress(status.progress || 0, status.status === "error" ? "聚类失败" : "聚类任务运行中", status.message || "");
    if (status.status === "done") {
      const result = await postJson("/api/cluster/result", { job_id: jobId });
      const scope = state.cluster[mode] || currentClusterState();
      scope.result = result;
      scope.detailPage = 1;
      scope.summaryPage = 1;
      scope.clusterFilter = "";
      if (state.cluster.mode === mode) {
        if (clusterDetailFilter) clusterDetailFilter.value = "";
        renderClusterResult();
      }
      showClusterProgress(100, "聚类完成", "结果已生成");
      window.setTimeout(hideClusterProgress, 900);
      showToast("地址聚类完成");
      return;
    }
    if (status.status === "error") {
      if (status.error_detail) console.error(status.error_detail);
      showToast(status.error || status.message || "聚类失败");
      return;
    }
    clusterPollTimer = window.setTimeout(() => pollClusterJob(jobId, mode), 900);
  } catch (error) {
    showToast(error.message || String(error));
  }
}

function showClusterProgress(percent, title, detail) {
  if (!clusterProgress) return;
  clusterProgress.hidden = false;
  const value = Math.max(0, Math.min(100, Number(percent) || 0));
  clusterProgressTitle.textContent = title;
  clusterProgressText.textContent = detail;
  clusterProgressBar.style.width = `${value}%`;
  clusterProgress.querySelector(".progress-track")?.setAttribute("aria-valuenow", String(Math.round(value)));
}

function hideClusterProgress() {
  if (!clusterProgress) return;
  clusterProgress.hidden = true;
  clusterProgressBar.style.width = "0%";
}

function renderClusterResult() {
  renderClusterSummaryCards();
  renderClusterDetailTable();
  renderClusterSummaryTable();
}

function renderClusterSummaryCards() {
  const container = document.querySelector("#clusterSummary");
  const result = currentClusterState().result;
  if (!container) return;
  if (!result) {
    container.innerHTML = "";
    return;
  }
  const stats = result.stats || {};
  const params = result.params || {};
  const items = [
    ["模式", params.mode === "incremental" ? "增量" : "全量"],
    ["样本数", formatNumber(stats.row_count)],
    ["簇数量", formatNumber(stats.cluster_count)],
    ["噪声数", formatNumber(stats.noise_count || 0)],
    ["模型", params.model || "-"],
    ["阈值/eps", params.threshold ?? params.eps ?? "-"],
  ];
  container.innerHTML = items.map(([label, value]) => `
    <div class="summary-card">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `).join("");
}

function clusterRows(kind) {
  const scope = currentClusterState();
  const table = kind === "summary" ? scope.result?.summary_table : scope.result?.detail_table;
  const rows = [...(table?.rows || [])];
  const filter = String(scope.clusterFilter || "").trim();
  if (kind === "detail" && filter) {
    return rows.filter((row) => String(row.addr_cluster_id).includes(filter));
  }
  return rows;
}

function sortedClusterRows(kind) {
  const rows = clusterRows(kind);
  const scope = currentClusterState();
  const sort = kind === "summary" ? scope.summarySort : scope.detailSort;
  const direction = sort.direction === "desc" ? -1 : 1;
  return rows.sort((left, right) => compareValues(left[sort.column], right[sort.column]) * direction);
}

function compareValues(left, right) {
  const leftNumber = Number(left);
  const rightNumber = Number(right);
  if (!Number.isNaN(leftNumber) && !Number.isNaN(rightNumber)) return leftNumber - rightNumber;
  return String(left ?? "").localeCompare(String(right ?? ""), "zh-CN");
}

function renderClusterDetailTable() {
  renderPagedClusterTable({
    kind: "detail",
    containerSelector: "#clusterDetailTable",
    paginationSelector: "#clusterDetailPagination",
    subtitleSelector: "#clusterDetailSubtitle",
    pageKey: "detailPage",
    sortKey: "detailSort",
    emptyText: "暂无聚类明细",
  });
}

function renderClusterSummaryTable() {
  renderPagedClusterTable({
    kind: "summary",
    containerSelector: "#clusterSummaryTable",
    paginationSelector: "#clusterSummaryPagination",
    subtitleSelector: "#clusterSummarySubtitle",
    pageKey: "summaryPage",
    sortKey: "summarySort",
    emptyText: "暂无聚类汇总",
  });
}

function renderPagedClusterTable(config) {
  const container = document.querySelector(config.containerSelector);
  const pagination = document.querySelector(config.paginationSelector);
  const subtitle = document.querySelector(config.subtitleSelector);
  const scope = currentClusterState();
  const table = config.kind === "summary" ? scope.result?.summary_table : scope.result?.detail_table;
  if (!container || !pagination || !subtitle) return;
  const columns = table?.columns || [];
  const rows = sortedClusterRows(config.kind);
  const pageSize = state.cluster.pageSize;
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  const page = Math.min(Math.max(1, Number(scope[config.pageKey] || 1)), totalPages);
  scope[config.pageKey] = page;
  subtitle.textContent = rows.length
    ? `共 ${formatNumber(rows.length)} 行；当前第 ${formatNumber(page)} / ${formatNumber(totalPages)} 页。`
    : config.emptyText;
  if (!columns.length || !rows.length) {
    container.innerHTML = config.emptyText;
    container.classList.add("empty");
    pagination.innerHTML = "";
    return;
  }
  container.classList.remove("empty");
  const sort = scope[config.sortKey];
  const start = (page - 1) * pageSize;
  const pageRows = rows.slice(start, start + pageSize);
  container.innerHTML = `
    <table>
      <thead>
        <tr>
          ${columns.map((column) => `
            <th><button class="sortable-th" type="button" data-column="${escapeHtml(column)}">${escapeHtml(column)}${sort.column === column ? (sort.direction === "asc" ? " ↑" : " ↓") : ""}</button></th>
          `).join("")}
        </tr>
      </thead>
      <tbody>
        ${pageRows.map((row) => `<tr>${columns.map((column) => `<td>${formatCell(row[column])}</td>`).join("")}</tr>`).join("")}
      </tbody>
    </table>
  `;
  container.querySelectorAll(".sortable-th").forEach((button) => {
    button.addEventListener("click", () => {
      const column = button.dataset.column;
      const current = scope[config.sortKey];
      scope[config.sortKey] = {
        column,
        direction: current.column === column && current.direction === "asc" ? "desc" : "asc",
      };
      scope[config.pageKey] = 1;
      renderPagedClusterTable(config);
    });
  });
  pagination.innerHTML = `
    <button type="button" data-page="prev" ${page <= 1 ? "disabled" : ""}>上一页</button>
    <span>第 ${formatNumber(page)} / ${formatNumber(totalPages)} 页</span>
    <label class="page-jump">跳至 <input type="number" min="1" max="${totalPages}" value="${page}" aria-label="跳转页码" /> 页</label>
    <button type="button" data-page="next" ${page >= totalPages ? "disabled" : ""}>下一页</button>
  `;
  pagination.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", () => {
      scope[config.pageKey] += button.dataset.page === "prev" ? -1 : 1;
      renderPagedClusterTable(config);
    });
  });
  bindPaginationJump(pagination, totalPages, (targetPage) => {
    scope[config.pageKey] = targetPage;
    renderPagedClusterTable(config);
  });
}

function bindPaginationJump(container, totalPages, onJump) {
  const input = container.querySelector(".page-jump input");
  if (!input) return;
  input.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    const rawPage = Number(input.value || 1);
    const targetPage = Math.min(Math.max(1, Math.floor(rawPage || 1)), totalPages);
    input.value = String(targetPage);
    onJump(targetPage);
  });
}

function exportClusterTable(kind, format) {
  const result = currentClusterState().result;
  const table = kind === "summary" ? result?.summary_table : result?.detail_table;
  if (!table?.columns?.length || !table?.rows?.length) {
    showToast("暂无可导出的聚类结果");
    return;
  }
  const rows = sortedClusterRows(kind);
  const name = kind === "summary" ? "cluster_summary" : "cluster_detail";
  exportRows(`${name}_${timestamp()}`, table.columns, rows, format);
}

function exportRows(filenameBase, columns, rows, format) {
  if (format === "xls") {
    const html = `
      <html><head><meta charset="utf-8" /></head><body>
        <table>
          <thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr></thead>
          <tbody>${rows.map((row) => `<tr>${columns.map((column) => `<td>${escapeHtml(row[column] ?? "")}</td>`).join("")}</tr>`).join("")}</tbody>
        </table>
      </body></html>
    `;
    downloadBlob(`${filenameBase}.xls`, html, "application/vnd.ms-excel;charset=utf-8");
    return;
  }
  const csv = "\ufeff" + [
    columns.map(csvEscape).join(","),
    ...rows.map((row) => columns.map((column) => csvEscape(row[column])).join(",")),
  ].join("\n");
  downloadBlob(`${filenameBase}.csv`, csv, "text/csv;charset=utf-8");
}

function exportClusterEngineeringFile() {
  const payload = currentClusterState().result?.engineering_file;
  if (!payload) {
    showToast("暂无可导出的工程文件");
    return;
  }
  downloadBlob(`cluster_engineering_${timestamp()}.json`, JSON.stringify(payload), "application/json;charset=utf-8");
}

function renderCommunityTable(rows) {
  const container = document.querySelector("#communityTable");
  if (!container) return;
  if (!rows?.length) {
    container.innerHTML = "<div class=\"empty\">暂无团伙</div>";
    return;
  }
  container.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>团伙</th>
          <th>类型</th>
          <th>中介</th>
          <th>借款人</th>
          <th>坏账率</th>
          <th>密度</th>
          <th>风险分</th>
          <th>中心中介</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map((row) => `
          <tr data-type="agent" data-id="${row.top_agents?.[0]?.agent_id || ""}">
            <td>${row.community_id}</td>
            <td>${row.community_type === "user" ? "用户社区" : "中介社区"}</td>
            <td>${formatNumber(row.agent_count)}</td>
            <td>${formatNumber(row.borrower_count)}</td>
            <td>${formatPercent(row.bad_debt_rate)}</td>
            <td>${formatPercent(row.density)}</td>
            <td>${formatNumber(row.risk_score)}</td>
            <td>${row.top_agents?.[0]?.label || "-"}</td>
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
  bindTableRows("#communityTable");
}

function renderAgentTable(rows) {
  document.querySelector("#agentTable").innerHTML = tableHtml(
    ["中介", "一度借款人", "二度中介数"],
    rows,
    (row) => [row.label, row.first_degree_borrowers, row.second_degree_agents],
    "agent",
    "agent_id",
  );
  bindTableRows("#agentTable");
}

function renderBorrowerTable(rows) {
  document.querySelector("#borrowerTable").innerHTML = tableHtml(
    ["借款人", "一度中介", "二度借款人", "二度逾期率"],
    rows,
    (row) => [row.label, row.first_degree_agents, row.second_degree_borrowers, formatPercent(row.second_degree_overdue_rate)],
    "borrower",
    "borrower_id",
  );
  bindTableRows("#borrowerTable");
}

function tableHtml(headers, rows, mapper, type, idKey) {
  if (!rows?.length) return "<div class=\"empty\">暂无数据</div>";
  return `
    <table>
      <thead><tr>${headers.map((item) => `<th>${item}</th>`).join("")}</tr></thead>
      <tbody>
        ${rows.map((row) => `
          <tr data-type="${type}" data-id="${row[idKey]}">
            ${mapper(row).map((value) => `<td>${value}</td>`).join("")}
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function bindTableRows(selector) {
  document.querySelectorAll(`${selector} tbody tr`).forEach((row) => {
    row.addEventListener("click", () => {
      setCenterType(row.dataset.type);
      centerInput.value = row.dataset.id;
      queryGraph();
    });
  });
}

function renderGraph(graph) {
  state.currentGraph = graph;
  const container = document.querySelector("#graphCanvas");
  const subtitle = document.querySelector("#graphSubtitle");
  const nodes = graph.nodes || [];
  const edges = graph.edges || [];
  renderLegendControls();
  if (!nodes.length) {
    container.innerHTML = "<div class=\"empty\">上传数据后展示图谱</div>";
    subtitle.textContent = "等待上传数据";
    return;
  }

  const width = Math.max(920, container.clientWidth || 920);
  const height = Math.max(560, container.clientHeight || 560);
  const visibleNodes = nodes.filter((node) => !graphHiddenTypes.has(node.type));
  const visibleNodeIds = new Set(visibleNodes.map((node) => node.id));
  const visibleEdges = edges.filter((edge) => visibleNodeIds.has(edge.source) && visibleNodeIds.has(edge.target));
  const positioned = positionNodes(visibleNodes, width, height);
  const byId = new Map(positioned.map((node) => [node.id, node]));
  const centerLabel = graph.center.type === "agent" ? "中介" : "借款人";
  subtitle.textContent = `${centerLabel} ${mask(graph.center.id)}，节点 ${visibleNodes.length}/${nodes.length}，边 ${visibleEdges.length}/${edges.length}`;

  container.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="二度关系图谱">
      <g class="graph-edges">
        ${visibleEdges.map((edge) => {
          const source = byId.get(edge.source);
          const target = byId.get(edge.target);
          if (!source || !target) return "";
          const relation = edge.relation || inferEdgeRelation(edge.source, edge.target);
          const meta = edgeMeta(relation);
          return `
            <line class="graph-edge" data-source="${edge.source}" data-target="${edge.target}" x1="${source.x}" y1="${source.y}" x2="${target.x}" y2="${target.y}" stroke="${meta.color}" stroke-width="${meta.width}">
              <title>${meta.label}</title>
            </line>
          `;
        }).join("")}
      </g>
      <g class="graph-nodes">
        ${positioned.map((node) => nodeSvg(node, graph.center.node_id)).join("")}
      </g>
    </svg>
  `;

  bindGraphViewport(container.querySelector("svg"), width, height);
  container.querySelector("svg")?.addEventListener("click", () => clearGraphHighlight(container));
  container.querySelectorAll(".graph-node").forEach((group) => {
    group.addEventListener("click", (event) => {
      event.stopPropagation();
      highlightNeighborhood(container, group.dataset.nodeId);
    });
    group.addEventListener("dblclick", (event) => {
      event.stopPropagation();
      setCenterType(group.dataset.type);
      centerInput.value = group.dataset.rawId;
      queryGraph();
    });
  });
}

function renderLegendControls() {
  const legend = document.querySelector(".legend");
  if (!legend) return;
  const items = [
    ["agent", "agent-dot", "中介"],
    ["borrower", "borrower-dot", "借款人"],
    ["device", "device-dot", "设备"],
    ["ip", "ip-dot", "IP"],
    ["address", "address-dot", "地址簇"],
  ];
  legend.innerHTML = `
    ${items.map(([type, dot, label]) => `
      <button class="legend-toggle ${graphHiddenTypes.has(type) ? "is-off" : ""}" type="button" data-type="${type}" aria-pressed="${!graphHiddenTypes.has(type)}">
        <i class="${dot}"></i>${label}
      </button>
    `).join("")}
    <span><i class="overdue-dot"></i>有逾期</span>
  `;
  legend.querySelectorAll(".legend-toggle").forEach((button) => {
    button.addEventListener("click", () => {
      const type = button.dataset.type;
      if (graphHiddenTypes.has(type)) graphHiddenTypes.delete(type);
      else graphHiddenTypes.add(type);
      if (state.currentGraph) renderGraph(state.currentGraph);
    });
  });
}

function highlightNeighborhood(container, nodeId) {
  const linked = new Set([nodeId]);
  container.querySelectorAll(".graph-edge").forEach((edge) => {
    if (edge.dataset.source === nodeId || edge.dataset.target === nodeId) {
      linked.add(edge.dataset.source);
      linked.add(edge.dataset.target);
      edge.classList.add("is-highlighted");
      edge.classList.remove("is-muted");
    } else {
      edge.classList.add("is-muted");
      edge.classList.remove("is-highlighted");
    }
  });
  container.querySelectorAll(".graph-node").forEach((node) => {
    node.classList.toggle("is-highlighted", linked.has(node.dataset.nodeId));
    node.classList.toggle("is-muted", !linked.has(node.dataset.nodeId));
  });
}

function clearGraphHighlight(container) {
  container.querySelectorAll(".is-muted, .is-highlighted").forEach((item) => {
    item.classList.remove("is-muted", "is-highlighted");
  });
}

function bindGraphViewport(svg, width, height) {
  if (!svg) return;
  const viewBox = { x: 0, y: 0, width, height };
  const apply = () => svg.setAttribute("viewBox", `${viewBox.x} ${viewBox.y} ${viewBox.width} ${viewBox.height}`);
  svg.addEventListener("click", (event) => {
    if (svg.dataset.dragged === "1") {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
  }, true);
  svg.addEventListener("wheel", (event) => {
    event.preventDefault();
    const scale = event.deltaY < 0 ? 0.86 : 1.16;
    const nextWidth = Math.max(width * 0.22, Math.min(width * 3, viewBox.width * scale));
    const nextHeight = Math.max(height * 0.22, Math.min(height * 3, viewBox.height * scale));
    const rect = svg.getBoundingClientRect();
    const px = (event.clientX - rect.left) / rect.width;
    const py = (event.clientY - rect.top) / rect.height;
    viewBox.x += (viewBox.width - nextWidth) * px;
    viewBox.y += (viewBox.height - nextHeight) * py;
    viewBox.width = nextWidth;
    viewBox.height = nextHeight;
    apply();
  }, { passive: false });

  let drag = null;
  svg.addEventListener("pointerdown", (event) => {
    drag = {
      x: event.clientX,
      y: event.clientY,
      startX: viewBox.x,
      startY: viewBox.y,
      moved: false,
    };
    svg.classList.add("is-panning");
    svg.setPointerCapture(event.pointerId);
  });
  svg.addEventListener("pointermove", (event) => {
    if (!drag) return;
    const rect = svg.getBoundingClientRect();
    if (Math.abs(event.clientX - drag.x) > 3 || Math.abs(event.clientY - drag.y) > 3) {
      drag.moved = true;
    }
    viewBox.x = drag.startX - ((event.clientX - drag.x) / rect.width) * viewBox.width;
    viewBox.y = drag.startY - ((event.clientY - drag.y) / rect.height) * viewBox.height;
    apply();
  });
  svg.addEventListener("pointerup", () => {
    if (drag?.moved) {
      svg.dataset.dragged = "1";
      window.setTimeout(() => {
        delete svg.dataset.dragged;
      }, 0);
    }
    drag = null;
    svg.classList.remove("is-panning");
  });
  svg.addEventListener("pointercancel", () => {
    drag = null;
    svg.classList.remove("is-panning");
  });
}

function inferEdgeRelation(source, target) {
  const sourceType = String(source).split(":", 1)[0];
  const targetType = String(target).split(":", 1)[0];
  return `${sourceType}_${targetType}`;
}

function edgeMeta(relation) {
  const normalized = relation === "borrower_agent" ? "agent_borrower" : relation;
  const meta = {
    agent_borrower: { color: "#8aa0ad", width: 1.35, label: "中介-用户" },
    borrower_device: { color: "#8b5bb5", width: 1.8, label: "用户-设备" },
    borrower_ip: { color: "#c78325", width: 1.6, label: "用户-IP" },
    agent_address: { color: "#596b78", width: 1.5, label: "中介-地址簇" },
  };
  return meta[normalized] || { color: "#b6c3cc", width: 1.2, label: "关联关系" };
}

function positionNodes(nodes, width, height) {
  const cx = width / 2;
  const cy = height / 2;
  const groups = {
    0: nodes.filter((node) => node.level === 0),
    1: nodes.filter((node) => node.level === 1),
    2: nodes.filter((node) => node.level >= 2),
  };
  const radiusOne = Math.min(width, height) * 0.23;
  const radiusTwo = Math.min(width, height) * 0.42;

  return nodes.map((node) => {
    if (node.level === 0) return { ...node, x: cx, y: cy };
    const ring = node.level === 1 ? groups[1] : groups[2];
    const index = ring.findIndex((item) => item.id === node.id);
    const count = Math.max(ring.length, 1);
    const radius = node.level === 1 ? radiusOne : radiusTwo;
    const angle = -Math.PI / 2 + (Math.PI * 2 * index) / count;
    const wobble = node.level === 1 ? 0 : (index % 5) * 6;
    return {
      ...node,
      x: cx + Math.cos(angle) * (radius + wobble),
      y: cy + Math.sin(angle) * (radius + wobble),
    };
  });
}

function nodeSvg(node, centerNodeId) {
  const isCenter = node.id === centerNodeId;
  const colors = {
    agent: ["var(--agent)", "var(--agent-soft)", "中"],
    borrower: ["var(--borrower)", "var(--borrower-soft)", "借"],
    device: ["#8b5bb5", "#eee5f6", "设"],
    ip: ["#c78325", "#f7ead7", "IP"],
    address: ["#596b78", "#e6ebef", "址"],
  };
  const [color, fill, mark] = colors[node.type] || ["#596b78", "#e6ebef", "?"];
  const stroke = node.overdue ? "var(--risk)" : color;
  const radius = isCenter ? 24 : node.level === 1 ? 18 : 14;
  const labelY = node.y + radius + 15;
  return `
    <g class="graph-node" data-node-id="${node.id}" data-type="${node.type}" data-raw-id="${node.raw_id}">
      <circle cx="${node.x}" cy="${node.y}" r="${radius}" fill="${fill}" stroke="${stroke}" stroke-width="${isCenter ? 4 : node.overdue ? 3 : 2}" />
      <text x="${node.x}" y="${node.y + 4}" text-anchor="middle" font-size="${isCenter ? 13 : 11}" fill="${color}" font-weight="700">
        ${mark}
      </text>
      <text class="graph-label" x="${node.x}" y="${labelY}" text-anchor="middle">${escapeHtml(node.label)}</text>
    </g>
  `;
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",", 2)[1]);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

async function postJson(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "请求失败");
  return data;
}

function formatNumber(value) {
  const number = Number(value || 0);
  return number.toLocaleString("zh-CN");
}

function formatFileSize(size) {
  const bytes = Number(size || 0);
  if (bytes >= 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
  if (bytes >= 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${bytes} B`;
}

function formatPercent(value) {
  const number = Number(value || 0);
  return `${(number * 100).toFixed(2)}%`;
}

function formatCell(value) {
  if (typeof value === "number") {
    return Number.isInteger(value) ? value : Number(value.toFixed(8));
  }
  return escapeHtml(value ?? "");
}

function mask(value) {
  const text = String(value || "");
  if (/^\d{7,}$/.test(text)) return `${text.slice(0, 3)}****${text.slice(-4)}`;
  if (text.length > 8) return `${text.slice(0, 4)}...${text.slice(-4)}`;
  return text;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function exportFeatureTable(format) {
  const { columns, rows } = state.featureTable || {};
  if (!columns?.length || !rows?.length) {
    showToast("暂无可导出的特征表");
    return;
  }
  if (format === "xls") {
    const html = `
      <html>
        <head><meta charset="utf-8" /></head>
        <body>
          <table>
            <thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr></thead>
            <tbody>
              ${rows.map((row) => `<tr>${columns.map((column) => `<td>${escapeHtml(row[column] ?? "")}</td>`).join("")}</tr>`).join("")}
            </tbody>
          </table>
        </body>
      </html>
    `;
    downloadBlob(`anti_fraud_features_${timestamp()}.xls`, html, "application/vnd.ms-excel;charset=utf-8");
    return;
  }
  const csv = "\ufeff" + [
    columns.map(csvEscape).join(","),
    ...rows.map((row) => columns.map((column) => csvEscape(row[column])).join(",")),
  ].join("\n");
  downloadBlob(`anti_fraud_features_${timestamp()}.csv`, csv, "text/csv;charset=utf-8");
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n\r]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function downloadBlob(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function timestamp() {
  const date = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}

function debounce(fn, wait) {
  let timer = 0;
  return (...args) => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => fn(...args), wait);
  };
}

function showToast(message) {
  const old = document.querySelector(".toast");
  if (old) old.remove();
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  document.body.appendChild(toast);
  window.setTimeout(() => toast.remove(), 2600);
}

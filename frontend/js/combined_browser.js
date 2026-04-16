let combinedGridApi = null;

const combinedState = {
  filePath: "",
  page: 1,
  pageSize: 200,
  totalRows: 0,
  columns: [],
  selectedProjects: [],
  selectedOrigins: [],
  selectedDepts: [],
  selectedTitles: [],
  runningOnly: true,
  nameSearch: "",
  viewMode: "general",
  showFee: false,
};

const GENERAL_COLS = [
  "Actual Name",
  "General Department",
  "Actual Title",
  "Actual Phone",
  "Actual Email",
  "Project",
  "GCMID",
  "Origin",
];

const PROJECT_COLS = [
  "Crew list name",
  "Project department",
  "Project job title",
  "Mobile number",
  "Crew email",
  "Project",
  "GCMID",
  "Origin",
];

function selectedValues(selectEl) {
  return Array.from(selectEl.selectedOptions).map((opt) => opt.value);
}

function setCombinedMessage(text, isError = false) {
  const el = document.querySelector("#combined-result-message");
  el.textContent = text;
  el.style.color = isError ? "#b00020" : "#1f6d2a";
}

function updateCombinedPager() {
  const totalPages = Math.max(1, Math.ceil((combinedState.totalRows || 0) / combinedState.pageSize));
  document.querySelector("#combined-page-info").textContent = `Page ${combinedState.page} / ${totalPages}`;
  document.querySelector("#combined-total-info").textContent = `${combinedState.totalRows} rows total`;
  document.querySelector("#combined-prev").disabled = combinedState.page <= 1;
  document.querySelector("#combined-next").disabled = combinedState.page >= totalPages;
}

function selectColumns(allColumns) {
  let selected;
  if (combinedState.viewMode === "general") {
    selected = GENERAL_COLS.filter((c) => allColumns.includes(c));
  } else if (combinedState.viewMode === "project") {
    selected = PROJECT_COLS.filter((c) => allColumns.includes(c));
  } else {
    selected = allColumns.slice();
  }

  if (combinedState.showFee && allColumns.includes("Daily fee") && !selected.includes("Daily fee")) {
    selected.push("Daily fee");
  }

  if (!combinedState.showFee) {
    selected = selected.filter((c) => c !== "Daily fee");
  }

  if (!selected.length) {
    return allColumns.slice(0, 30);
  }
  return selected;
}

function updateCombinedColumns(columns) {
  const chosen = selectColumns(columns);
  const defs = chosen.map((col) => {
    const def = { field: col, headerName: col, sortable: true, filter: true, resizable: true };
    if (col === "Daily fee") {
      def.valueFormatter = (params) => {
        const raw = (params.value ?? "").toString().trim();
        if (!raw) return "";
        const normalized = raw.replace(/[^0-9.-]/g, "");
        const num = Number(normalized);
        if (!Number.isFinite(num)) return raw;
        return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(Math.round(num));
      };
    }
    return def;
  });
  combinedGridApi.setGridOption("columnDefs", defs);
}

function buildFilters() {
  const filters = [];
  if (combinedState.runningOnly) {
    const today = new Date().toISOString().slice(0, 10);
    filters.push({ col: "Project end date", op: "gte_date", value: today });
  }
  if (combinedState.selectedProjects.length) {
    filters.push({ col: "Project", op: "in", value: combinedState.selectedProjects });
  }
  if (combinedState.selectedOrigins.length) {
    filters.push({ col: "Origin", op: "in", value: combinedState.selectedOrigins });
  }
  if (combinedState.selectedDepts.length) {
    filters.push({ col: "General Department", op: "in", value: combinedState.selectedDepts });
  }
  if (combinedState.selectedTitles.length) {
    filters.push({ col: "General Title", op: "in", value: combinedState.selectedTitles });
  }
  if ((combinedState.nameSearch || "").trim()) {
    filters.push({ col: "Crew list name", op: "contains_normalized", value: combinedState.nameSearch.trim() });
  }
  return filters;
}

async function initCombinedFile() {
  const files = await browseFiles();
  const file = (files.combined || [])[0];
  if (!file) {
    setCombinedMessage("Combined file not found.", true);
    return false;
  }
  combinedState.filePath = file.path;
  document.querySelector("#combined-file-info").textContent = `${file.name} (${file.mtime})`;
  return true;
}

function fillMultiSelect(selectId, values, selected = []) {
  const select = document.querySelector(selectId);
  const selectedSet = new Set(selected);
  select.innerHTML = "";
  for (const value of values) {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    if (selectedSet.has(value)) opt.selected = true;
    select.appendChild(opt);
  }
}

async function loadFilterOptions() {
  const baseFilters = [];
  if (combinedState.runningOnly) {
    baseFilters.push({ col: "Project end date", op: "gte_date", value: new Date().toISOString().slice(0, 10) });
  }

  const projects = await browseDistinct({
    file_path: combinedState.filePath,
    sheet: "Combined",
    columns: ["Project"],
    filters: baseFilters,
    limit: 5000,
  });
  const projectValues = projects.rows.map((r) => r.Project).filter(Boolean);
  fillMultiSelect("#combined-project-filter", projectValues, combinedState.selectedProjects);

  const origins = await browseDistinct({
    file_path: combinedState.filePath,
    sheet: "Combined",
    columns: ["Origin"],
    filters: [],
    limit: 1000,
  });
  const originValues = origins.rows.map((r) => r.Origin).filter(Boolean);
  fillMultiSelect("#combined-origin-filter", originValues, combinedState.selectedOrigins);

  const deptRows = await browseDistinct({
    file_path: combinedState.filePath,
    sheet: "Combined",
    columns: ["General Department", "Department ID"],
    filters: [],
    limit: 5000,
  });
  const sortedDepts = deptRows.rows
    .filter((r) => r["General Department"])
    .sort((a, b) => Number(a["Department ID"] || 999999) - Number(b["Department ID"] || 999999))
    .map((r) => r["General Department"]);
  fillMultiSelect("#combined-dept-filter", [...new Set(sortedDepts)], combinedState.selectedDepts);

  const titleFilters = [];
  if (combinedState.selectedDepts.length) {
    titleFilters.push({ col: "General Department", op: "in", value: combinedState.selectedDepts });
  }
  const titleRows = await browseDistinct({
    file_path: combinedState.filePath,
    sheet: "Combined",
    columns: ["General Title", "Title ID"],
    filters: titleFilters,
    limit: 5000,
  });
  const sortedTitles = titleRows.rows
    .filter((r) => r["General Title"])
    .sort((a, b) => Number(a["Title ID"] || 999999) - Number(b["Title ID"] || 999999))
    .map((r) => r["General Title"]);
  fillMultiSelect("#combined-title-filter", [...new Set(sortedTitles)], combinedState.selectedTitles);
}

async function loadCombinedPage() {
  if (!combinedState.filePath) return;

  const payload = await browseQuery({
    file_path: combinedState.filePath,
    sheet: "Combined",
    page: combinedState.page,
    page_size: combinedState.pageSize,
    sort: [
      { col: "Project start date", dir: "desc" },
      { col: "Department ID", dir: "asc" },
      { col: "Title ID", dir: "asc" },
    ],
    filters: buildFilters(),
  });

  combinedState.columns = payload.columns || [];
  updateCombinedColumns(combinedState.columns);
  combinedGridApi.setGridOption("rowData", payload.rows || []);
  combinedState.totalRows = payload.total_rows || 0;
  updateCombinedPager();
  setCombinedMessage(`Loaded ${payload.rows.length} rows.`);
}

async function exportToExcel() {
  if (!combinedState.filePath) return;
  const today = new Date().toISOString().slice(0, 10);
  const payload = {
    file_path: combinedState.filePath,
    sheet: "Combined",
    sort: [
      { col: "Project start date", dir: "desc" },
      { col: "Department ID", dir: "asc" },
      { col: "Title ID", dir: "asc" },
    ],
    filters: buildFilters(),
    filename: `CrewIndex_Combined_${today}`,
  };
  setCombinedMessage("Preparing export...");
  try {
    await browseExport(payload);
    setCombinedMessage("Export downloaded.");
  } catch (err) {
    setCombinedMessage(err.message || "Export failed.", true);
  }
}

async function refreshAll() {
  if (!(await initCombinedFile())) return;
  await loadFilterOptions();
  await loadCombinedPage();
}

function wireFilterChange(selectId, targetField) {
  document.querySelector(selectId).addEventListener("change", async (e) => {
    combinedState[targetField] = selectedValues(e.target);
    if (selectId === "#combined-dept-filter") {
      combinedState.selectedTitles = [];
      await loadFilterOptions();
    }
    combinedState.page = 1;
    await loadCombinedPage();
  });
}

document.addEventListener("DOMContentLoaded", async () => {
  combinedGridApi = agGrid.createGrid(document.querySelector("#combined-grid"), {
    columnDefs: [],
    rowData: [],
    defaultColDef: { sortable: true, filter: true, resizable: true },
  });

  document.querySelector("#combined-page-size").addEventListener("change", async (e) => {
    combinedState.pageSize = Number(e.target.value || 200);
    combinedState.page = 1;
    await loadCombinedPage();
  });

  document.querySelector("#combined-prev").addEventListener("click", async () => {
    if (combinedState.page > 1) {
      combinedState.page -= 1;
      await loadCombinedPage();
    }
  });

  document.querySelector("#combined-next").addEventListener("click", async () => {
    combinedState.page += 1;
    await loadCombinedPage();
  });

  document.querySelector("#combined-refresh").addEventListener("click", refreshAll);

  document.querySelector("#combined-search").addEventListener("input", (e) => {
    combinedState.nameSearch = e.target.value || "";
    combinedState.page = 1;
    loadCombinedPage();
  });

  document.querySelector("#combined-export-csv").addEventListener("click", () => {
    combinedGridApi.exportDataAsCsv({ fileName: `combined_browser_page_${combinedState.page}.csv` });
  });

  document.querySelector("#export-excel").addEventListener("click", exportToExcel);

  document.querySelector("#combined-running-only").addEventListener("change", async (e) => {
    combinedState.runningOnly = Boolean(e.target.checked);
    combinedState.page = 1;
    await loadFilterOptions();
    await loadCombinedPage();
  });

  wireFilterChange("#combined-project-filter", "selectedProjects");
  wireFilterChange("#combined-origin-filter", "selectedOrigins");
  wireFilterChange("#combined-dept-filter", "selectedDepts");
  wireFilterChange("#combined-title-filter", "selectedTitles");

  document.querySelectorAll('input[name="combined-view-mode"]').forEach((radio) => {
    radio.addEventListener("change", () => {
      combinedState.viewMode = radio.value;
      updateCombinedColumns(combinedState.columns);
    });
  });

  document.querySelector("#combined-show-fee").addEventListener("change", (e) => {
    combinedState.showFee = Boolean(e.target.checked);
    updateCombinedColumns(combinedState.columns);
  });

  await refreshAll();
});

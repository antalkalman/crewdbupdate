let crewGridApi = null;

const state = {
  filePath: "",
  page: 1,
  pageSize: 200,
  totalRows: 0,
  columns: [],
  activeOnly: false,
  viewMode: "general",
  showFee: false,
};

const filters = {
  name: "",
  titles: [],
  departments: [],
  projects: [],
  statuses: [],
  origins: [],
};

const OPTIONS = {
  title: [],
  dept: [],
  project: [],
  status: ["Active", "Retired", "Foreign", "External"],
  origin: ["SFlist", "Historical"],
};

const DEPT_TITLE_MAP = {};  // dept name → array of title names

const GENERAL_COLS = [
  "Actual Name", "General Department", "Actual Title",
  "Last Phone", "Last Email", "Project", "GCMID",
  "Origin", "Status",
];

const PROJECT_COLS = [
  "Crew list name", "Project department", "Project job title",
  "Phone", "Email", "Project", "General Title",
  "GCMID", "Origin", "State",
];

const FEE_COLS = ["Daily fee", "Weekly fee", "Deal type", "Business type"];

const EXPORT_COLS = [
  "CM ID", "Actual Name", "Actual Title", "Last General Title",
  "Last Department", "Actual Phone", "Actual Email",
  "Last Phone", "Last Email", "Status", "Shows Worked",
];

// ── Utilities ────────────────────────────────────────────────────────────────

function formatPhone(val) {
  if (!val) return "";
  const s = String(val).replace(/\D/g, "");
  if (s.length < 8) return s;
  return `+${s.slice(0, 2)} ${s.slice(2, 4)} ${s.slice(4, 8)} ${s.slice(8)}`.trim();
}

function setMessage(text, isError) {
  const el = document.querySelector("#crew-message");
  el.textContent = text;
  el.style.color = isError ? "#b00020" : "";
}

// ── Column selection ─────────────────────────────────────────────────────────

function selectColumns(allColumns) {
  let selected;
  if (state.viewMode === "general") {
    selected = GENERAL_COLS.filter((c) => allColumns.includes(c));
  } else if (state.viewMode === "project") {
    selected = PROJECT_COLS.filter((c) => allColumns.includes(c));
  } else {
    selected = allColumns.slice();
  }

  if (state.showFee) {
    FEE_COLS.forEach((c) => {
      if (allColumns.includes(c) && !selected.includes(c)) selected.push(c);
    });
  } else {
    selected = selected.filter((c) => !FEE_COLS.includes(c));
  }

  selected = selected.filter((c) => c !== "Department ID" && c !== "Title ID");
  return selected.length ? selected : allColumns.slice(0, 30);
}

function buildColumnDefs(columns) {
  const chosen = selectColumns(columns);
  return chosen.map((col) => {
    const def = { field: col, headerName: col, sortable: true, filter: true, resizable: true };

    if (col === "GCMID") def.width = 80;
    else if (["Actual Name", "Crew list name"].includes(col)) def.width = 180;
    else if (["General Department", "Project department"].includes(col)) def.width = 170;
    else if (["Actual Title", "General Title", "Project job title"].includes(col)) def.width = 200;
    else if (["Last Email", "Email"].includes(col)) def.width = 220;
    else if (["Last Phone", "Phone"].includes(col)) {
      def.width = 160;
      def.valueFormatter = (p) => formatPhone(p.value);
    }
    else if (col === "Project") def.width = 150;
    else if (col === "Origin") def.width = 90;
    else if (col === "Status") {
      def.width = 100;
      def.cellStyle = (p) => {
        const v = p.value;
        if (v === "Active") return { color: "#1f6d2a", fontWeight: "600" };
        if (v === "Retired") return { color: "#888" };
        if (v === "Foreign") return { color: "#1a5fa8" };
        if (v === "External") return { color: "#b00020" };
        return {};
      };
    }
    else if (["Daily fee", "Weekly fee"].includes(col)) {
      def.width = 110;
      def.valueFormatter = (p) => {
        const raw = (p.value ?? "").toString().trim();
        if (!raw) return "";
        const num = Number(raw.replace(/[^0-9.-]/g, ""));
        return Number.isFinite(num) ? new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(Math.round(num)) : raw;
      };
    }
    else if (col === "State") def.width = 90;
    else def.width = 130;

    return def;
  });
}

function updateColumns() {
  crewGridApi.setGridOption("columnDefs", buildColumnDefs(state.columns));
}

// ── Filter building ──────────────────────────────────────────────────────────

function buildFilters() {
  const f = [];
  if (state.activeOnly) {
    const today = new Date().toISOString().slice(0, 10);
    f.push({ col: "Project end date", op: "gte_date", value: today });
  }
  if (filters.name.trim()) {
    f.push({ col: "Crew list name", op: "contains_normalized", value: filters.name.trim() });
  }
  if (filters.titles.length) {
    f.push({ col: "General Title|Actual Title", op: "cross_column_in", value: filters.titles });
  }
  if (filters.departments.length) {
    f.push({ col: "General Department", op: "in", value: filters.departments });
  }
  if (filters.projects.length) {
    f.push({ col: "Project", op: "in", value: filters.projects });
  }
  if (filters.statuses.length) {
    f.push({ col: "Status", op: "in", value: filters.statuses });
  }
  if (filters.origins.length) {
    f.push({ col: "Origin", op: "in", value: filters.origins });
  }
  return f;
}

// ── Pagination ───────────────────────────────────────────────────────────────

function updatePager() {
  const totalPages = Math.max(1, Math.ceil(state.totalRows / state.pageSize));
  document.querySelector("#crew-page-info").textContent = `Page ${state.page} / ${totalPages}`;
  document.querySelector("#crew-total-info").textContent = `${state.totalRows.toLocaleString()} rows total`;
  document.querySelector("#crew-prev").disabled = state.page <= 1;
  document.querySelector("#crew-next").disabled = state.page >= totalPages;
}

function updateRowCount() {
  const displayed = crewGridApi.getDisplayedRowCount();
  document.querySelector("#crew-row-count").textContent =
    `${displayed.toLocaleString()} of ${state.totalRows.toLocaleString()} rows`;
}

// ── Data loading ─────────────────────────────────────────────────────────────

async function initCrewFile() {
  const files = await browseFiles();
  const file = (files.crewindex || [])[0];
  if (!file) {
    setMessage("CrewIndex.xlsx not found.", true);
    return false;
  }
  state.filePath = file.path;
  return true;
}

async function loadPage() {
  if (state.viewMode === "export") {
    await loadRegistryPage();
    return;
  }
  if (!state.filePath) return;
  setMessage("Loading...");

  try {
    const payload = await browseQuery({
      file_path: state.filePath,
      sheet: "CrewIndex",
      page: state.page,
      page_size: state.pageSize,
      sort: [
        { col: "Project start date", dir: "desc" },
        { col: "Department ID", dir: "asc" },
        { col: "Title ID", dir: "asc" },
      ],
      filters: buildFilters(),
    });

    state.columns = payload.columns || [];
    updateColumns();
    crewGridApi.setGridOption("rowData", payload.rows || []);
    state.totalRows = payload.total_rows || 0;
    updatePager();
    updateRowCount();
    setMessage(`Page ${state.page} loaded.`);
  } catch (err) {
    setMessage(err.message || "Failed to load data.", true);
  }
}

function triggerReload() {
  state.page = 1;
  loadPage();
}

// ── Export view (two-step: CrewIndex → GCMIDs → CrewRegistry) ────────────────

async function fetchMatchingGCMIDs() {
  try {
    const result = await browseDistinct({
      file_path: state.filePath,
      sheet: "CrewIndex",
      columns: ["GCMID"],
      filters: buildFilters(),
      limit: 5000,
    });
    return result.rows
      .map((r) => String(r["GCMID"] || "").split(".")[0].trim())
      .filter((v) => v && v !== "" && v !== "nan");
  } catch (e) {
    console.error("Failed to fetch GCMIDs:", e);
    return [];
  }
}

function buildExportColDefs(columns) {
  return columns.map((col) => {
    const def = { field: col, headerName: col, sortable: true, filter: true, resizable: true };
    if (col === "CM ID") def.width = 70;
    else if (col === "Shows Worked") { def.flex = 2; }
    else if (col === "Actual Name") { def.width = 180; }
    else if (["Actual Phone", "Last Phone"].includes(col)) {
      def.width = 150;
      def.valueFormatter = (p) => formatPhone(p.value);
    }
    else if (["Actual Email", "Last Email"].includes(col)) def.width = 210;
    else if (col === "Status") {
      def.width = 100;
      def.cellStyle = (p) => {
        const v = p.value;
        if (v === "Active") return { color: "#1f6d2a", fontWeight: "600" };
        if (v === "Retired") return { color: "#888" };
        if (v === "Foreign") return { color: "#1a5fa8" };
        if (v === "External") return { color: "#b00020" };
        return {};
      };
    }
    else def.width = 170;
    return def;
  });
}

async function loadRegistryPage() {
  setMessage("Finding matching crew...");
  try {
    const gcmids = await fetchMatchingGCMIDs();
    if (gcmids.length === 0) {
      crewGridApi.setGridOption("columnDefs", buildExportColDefs(EXPORT_COLS));
      crewGridApi.setGridOption("rowData", []);
      state.totalRows = 0;
      document.querySelector("#crew-page-info").textContent = "";
      document.querySelector("#crew-total-info").textContent = "";
      document.querySelector("#crew-prev").style.display = "none";
      document.querySelector("#crew-next").style.display = "none";
      document.querySelector("#crew-row-count").textContent = "0 people";
      setMessage("No matching crew found.");
      return;
    }

    const result = await browseRegistryLookup(gcmids);
    crewGridApi.setGridOption("columnDefs", buildExportColDefs(result.columns || EXPORT_COLS));
    crewGridApi.setGridOption("rowData", result.rows || []);
    state.totalRows = result.total_rows || 0;

    document.querySelector("#crew-page-info").textContent = "";
    document.querySelector("#crew-total-info").textContent = "";
    document.querySelector("#crew-prev").style.display = "none";
    document.querySelector("#crew-next").style.display = "none";
    document.querySelector("#crew-row-count").textContent =
      `${state.totalRows.toLocaleString()} people`;
    setMessage(`${state.totalRows.toLocaleString()} crew members found.`);
  } catch (err) {
    setMessage(err.message || "Failed to load registry.", true);
  }
}

async function exportToExcel() {
  const today = new Date().toISOString().slice(0, 10);
  setMessage("Preparing export...");
  try {
    if (state.viewMode === "export") {
      const gcmids = await fetchMatchingGCMIDs();
      if (!gcmids.length) {
        setMessage("No matching crew to export.", true);
        return;
      }
      await browseRegistryExport(gcmids, `CrewExport_${today}`);
    } else {
      if (!state.filePath) return;
      await browseExport({
        file_path: state.filePath,
        sheet: "CrewIndex",
        sort: [
          { col: "Project start date", dir: "desc" },
          { col: "Department ID", dir: "asc" },
          { col: "Title ID", dir: "asc" },
        ],
        filters: buildFilters(),
        filename: `CrewExplorer_${today}`,
      });
    }
    setMessage("Export downloaded.");
  } catch (err) {
    setMessage(err.message || "Export failed.", true);
  }
}

// ── Generic tag picker factory ───────────────────────────────────────────────

function createPicker(inputId, dropdownId, tagsId, getOptions, filterKey, onChange) {
  const input = document.querySelector(inputId);
  const dropdown = document.querySelector(dropdownId);
  const tagsEl = document.querySelector(tagsId);

  function renderTags() {
    tagsEl.innerHTML = "";
    filters[filterKey].forEach((val) => {
      const tag = document.createElement("span");
      tag.className = "filter-tag";
      const label = document.createTextNode(val + " ");
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = "\u00d7";
      btn.addEventListener("click", () => {
        filters[filterKey] = filters[filterKey].filter((v) => v !== val);
        renderTags();
        triggerReload();
        if (onChange) onChange();
      });
      tag.appendChild(label);
      tag.appendChild(btn);
      tagsEl.appendChild(tag);
    });
  }

  function renderDropdown(query) {
    const opts = getOptions();
    const filtered = opts
      .filter((o) => !filters[filterKey].includes(o))
      .filter((o) => o.toLowerCase().includes(query.toLowerCase()));
    if (!filtered.length) { dropdown.style.display = "none"; return; }
    dropdown.innerHTML = "";
    filtered.slice(0, 60).forEach((opt) => {
      const item = document.createElement("div");
      item.className = "filter-dropdown-item";
      item.textContent = opt;
      item.addEventListener("mousedown", (e) => {
        e.preventDefault();
        filters[filterKey] = [...filters[filterKey], opt];
        renderTags();
        input.value = "";
        dropdown.style.display = "none";
        triggerReload();
        if (onChange) onChange();
      });
      dropdown.appendChild(item);
    });
    dropdown.style.display = "block";
  }

  function refreshDropdown() {
    if (dropdown.style.display === "block") {
      renderDropdown(input.value);
    }
  }

  input.addEventListener("input", (e) => renderDropdown(e.target.value));
  input.addEventListener("focus", (e) => renderDropdown(e.target.value));
  input.addEventListener("blur", () => setTimeout(() => (dropdown.style.display = "none"), 150));

  return { renderTags, refreshDropdown };
}

// ── Load filter options from API ─────────────────────────────────────────────

async function loadFilterOptions() {
  try {
    const gt = await fetchGeneralTitles();
    OPTIONS.title = gt.titles || [];
  } catch (_) {}

  // Build dept→titles map
  try {
    const resp = await fetch("/api/general_titles_with_dept").then((r) => r.json());
    Object.keys(DEPT_TITLE_MAP).forEach((k) => delete DEPT_TITLE_MAP[k]);
    (resp.items || []).forEach(({ dept, title }) => {
      if (!DEPT_TITLE_MAP[dept]) DEPT_TITLE_MAP[dept] = [];
      DEPT_TITLE_MAP[dept].push(title);
    });
  } catch (_) {}

  if (!state.filePath) return;

  try {
    const depts = await browseDistinct({
      file_path: state.filePath,
      sheet: "CrewIndex",
      columns: ["General Department"],
      filters: [],
      limit: 500,
    });
    OPTIONS.dept = depts.rows.map((r) => r["General Department"]).filter(Boolean).sort();
  } catch (_) {}

  try {
    const projs = await browseDistinct({
      file_path: state.filePath,
      sheet: "CrewIndex",
      columns: ["Project", "Project start date"],
      filters: [],
      limit: 500,
    });
    OPTIONS.project = projs.rows
      .filter((r) => r["Project"])
      .sort((a, b) => (b["Project start date"] || "").localeCompare(a["Project start date"] || ""))
      .map((r) => r["Project"])
      .filter((v, i, arr) => arr.indexOf(v) === i);
  } catch (_) {}
}

// ── Refresh ──────────────────────────────────────────────────────────────────

async function refreshAll() {
  if (!(await initCrewFile())) return;
  await loadFilterOptions();
  await loadPage();
}

// ── Init ─────────────────────────────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", async () => {
  crewGridApi = agGrid.createGrid(document.querySelector("#crew-grid"), {
    columnDefs: [],
    rowData: [],
    defaultColDef: { sortable: true, filter: true, resizable: true },
    onFilterChanged: () => updateRowCount(),
  });

  // Pickers
  const pickers = {
    title:   createPicker("#picker-title",   "#dropdown-title",   "#tags-title",   () => {
      if (filters.departments.length > 0) {
        const allowed = new Set();
        filters.departments.forEach((d) => (DEPT_TITLE_MAP[d] || []).forEach((t) => allowed.add(t)));
        return OPTIONS.title.filter((t) => allowed.has(t));
      }
      return OPTIONS.title;
    }, "titles"),
    dept:    createPicker("#picker-dept",     "#dropdown-dept",    "#tags-dept",    () => OPTIONS.dept,    "departments", () => pickers.title.refreshDropdown()),
    project: createPicker("#picker-project", "#dropdown-project", "#tags-project", () => OPTIONS.project, "projects"),
    status:  createPicker("#picker-status",  "#dropdown-status",  "#tags-status",  () => OPTIONS.status,  "statuses"),
    origin:  createPicker("#picker-origin",  "#dropdown-origin",  "#tags-origin",  () => OPTIONS.origin,  "origins"),
  };

  // Name search — debounced
  let nameTimer = null;
  document.querySelector("#filter-name").addEventListener("input", (e) => {
    clearTimeout(nameTimer);
    nameTimer = setTimeout(() => {
      filters.name = e.target.value;
      triggerReload();
    }, 400);
  });

  // Active projects only
  document.querySelector("#crew-active-only").addEventListener("change", (e) => {
    state.activeOnly = e.target.checked;
    triggerReload();
  });

  // Show fees
  document.querySelector("#crew-show-fee").addEventListener("change", (e) => {
    state.showFee = Boolean(e.target.checked);
    updateColumns();
  });

  // Page size
  document.querySelector("#crew-page-size").addEventListener("change", (e) => {
    state.pageSize = Number(e.target.value || 200);
    triggerReload();
  });

  // Pagination
  document.querySelector("#crew-prev").addEventListener("click", () => {
    if (state.page > 1) { state.page--; loadPage(); }
  });
  document.querySelector("#crew-next").addEventListener("click", () => {
    state.page++;
    loadPage();
  });

  // View mode radios
  document.querySelectorAll('input[name="crew-view-mode"]').forEach((radio) => {
    radio.addEventListener("change", () => {
      const wasExport = state.viewMode === "export";
      state.viewMode = radio.value;
      state.page = 1;
      if (wasExport && radio.value !== "export") {
        document.querySelector("#crew-prev").style.display = "";
        document.querySelector("#crew-next").style.display = "";
      }
      if (radio.value === "export" || wasExport) {
        loadPage();
      } else {
        updateColumns();
        loadPage();
      }
    });
  });

  // Clear all filters
  document.querySelector("#crew-clear-filters").addEventListener("click", () => {
    filters.name = "";
    filters.titles = [];
    filters.departments = [];
    filters.projects = [];
    filters.statuses = [];
    filters.origins = [];
    document.querySelector("#filter-name").value = "";
    Object.values(pickers).forEach((p) => p.renderTags());
    triggerReload();
  });

  // Export
  document.querySelector("#crew-export-excel").addEventListener("click", exportToExcel);

  // Refresh
  document.querySelector("#crew-refresh").addEventListener("click", async () => {
    await loadFilterOptions();
    loadPage();
  });

  await refreshAll();
});

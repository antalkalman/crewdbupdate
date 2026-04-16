let namesGridApi = null;
let namesState = {
  filePath: "",
  page: 1,
  pageSize: 200,
  totalRows: 0,
  columns: [],
  search: "",
};

function setNamesMessage(text, isError = false) {
  const el = document.querySelector("#names-result-message");
  el.textContent = text;
  el.style.color = isError ? "#b00020" : "#1f6d2a";
}

function updateNamesPager() {
  const totalPages = Math.max(1, Math.ceil((namesState.totalRows || 0) / namesState.pageSize));
  document.querySelector("#names-page-info").textContent = `Page ${namesState.page} / ${totalPages}`;
  document.querySelector("#names-total-info").textContent = `${namesState.totalRows} rows total`;
  document.querySelector("#names-prev").disabled = namesState.page <= 1;
  document.querySelector("#names-next").disabled = namesState.page >= totalPages;
}

function updateNamesColumns(columns) {
  const defs = columns.map((col) => ({ field: col, headerName: col, sortable: true, filter: true, resizable: true }));
  namesGridApi.setGridOption("columnDefs", defs);
}

async function initNamesFile() {
  const files = await browseFiles();
  const file = (files.names || [])[0];
  if (!file) {
    setNamesMessage("Names file not found.", true);
    return false;
  }
  namesState.filePath = file.path;
  document.querySelector("#names-file-info").textContent = `${file.name} (${file.mtime})`;
  return true;
}

async function loadNamesPage() {
  if (!namesState.filePath) return;

  const filters = [];
  if ((namesState.search || "").trim()) {
    filters.push({ col: "*", op: "contains_any", value: namesState.search.trim() });
  }

  const payload = await browseQuery({
    file_path: namesState.filePath,
    sheet: "Names",
    page: namesState.page,
    page_size: namesState.pageSize,
    sort: [],
    filters,
  });

  if (!namesState.columns.length || JSON.stringify(namesState.columns) !== JSON.stringify(payload.columns)) {
    namesState.columns = payload.columns || [];
    updateNamesColumns(namesState.columns);
  }

  namesGridApi.setGridOption("rowData", payload.rows || []);
  namesState.totalRows = payload.total_rows || 0;
  updateNamesPager();
  setNamesMessage(`Loaded ${payload.rows.length} rows.`);
}

document.addEventListener("DOMContentLoaded", async () => {
  namesGridApi = agGrid.createGrid(document.querySelector("#names-grid"), {
    columnDefs: [],
    rowData: [],
    defaultColDef: { sortable: true, filter: true, resizable: true },
  });

  document.querySelector("#names-page-size").addEventListener("change", async (e) => {
    namesState.pageSize = Number(e.target.value || 200);
    namesState.page = 1;
    await loadNamesPage();
  });

  document.querySelector("#names-prev").addEventListener("click", async () => {
    if (namesState.page > 1) {
      namesState.page -= 1;
      await loadNamesPage();
    }
  });

  document.querySelector("#names-next").addEventListener("click", async () => {
    namesState.page += 1;
    await loadNamesPage();
  });

  document.querySelector("#names-refresh").addEventListener("click", async () => {
    if (await initNamesFile()) {
      await loadNamesPage();
    }
  });

  document.querySelector("#names-search").addEventListener("input", (e) => {
    namesState.search = e.target.value || "";
    namesState.page = 1;
    loadNamesPage();
  });

  document.querySelector("#names-export-csv").addEventListener("click", () => {
    namesGridApi.exportDataAsCsv({ fileName: `names_browser_page_${namesState.page}.csv` });
  });

  if (await initNamesFile()) {
    await loadNamesPage();
  }
});

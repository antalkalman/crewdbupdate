let sfGridApi = null;
let sfState = {
  filePath: "",
  page: 1,
  pageSize: 200,
  totalRows: 0,
  columns: [],
};

function setSfMessage(text, isError = false) {
  const el = document.querySelector("#sf-result-message");
  el.textContent = text;
  el.style.color = isError ? "#b00020" : "#1f6d2a";
}

function updateSfPager() {
  const totalPages = Math.max(1, Math.ceil((sfState.totalRows || 0) / sfState.pageSize));
  document.querySelector("#sf-page-info").textContent = `Page ${sfState.page} / ${totalPages}`;
  document.querySelector("#sf-total-info").textContent = `${sfState.totalRows} rows total`;
  document.querySelector("#sf-prev").disabled = sfState.page <= 1;
  document.querySelector("#sf-next").disabled = sfState.page >= totalPages;
}

function updateSfColumns(columns) {
  const defs = columns.map((col) => ({ field: col, headerName: col, sortable: true, filter: true, resizable: true }));
  sfGridApi.setGridOption("columnDefs", defs);
}

async function loadSfFiles() {
  const files = await browseFiles();
  const select = document.querySelector("#sf-file-select");
  select.innerHTML = "";
  for (const file of files.sf_exports || []) {
    const opt = document.createElement("option");
    opt.value = file.path;
    opt.textContent = `${file.name} (${file.mtime})`;
    select.appendChild(opt);
  }

  if (!select.value) {
    sfState.filePath = "";
    sfGridApi.setGridOption("rowData", []);
    setSfMessage("No SF export files found.", true);
    return;
  }

  sfState.filePath = select.value;
  sfState.page = 1;
}

async function loadSfPage() {
  if (!sfState.filePath) return;

  const payload = await browseQuery({
    file_path: sfState.filePath,
    sheet: "",
    page: sfState.page,
    page_size: sfState.pageSize,
    sort: [],
    filters: [],
  });

  if (!sfState.columns.length || JSON.stringify(sfState.columns) !== JSON.stringify(payload.columns)) {
    sfState.columns = payload.columns || [];
    updateSfColumns(sfState.columns);
  }

  sfGridApi.setGridOption("rowData", payload.rows || []);
  sfState.totalRows = payload.total_rows || 0;
  updateSfPager();
  setSfMessage(`Loaded ${payload.rows.length} rows from ${sfState.filePath}`);
}

async function refreshSfAll() {
  try {
    await loadSfFiles();
    await loadSfPage();
  } catch (error) {
    setSfMessage(error.message || "Failed to load SF browser data.", true);
  }
}

document.addEventListener("DOMContentLoaded", async () => {
  sfGridApi = agGrid.createGrid(document.querySelector("#sf-grid"), {
    columnDefs: [],
    rowData: [],
    defaultColDef: { sortable: true, filter: true, resizable: true },
  });

  document.querySelector("#sf-file-select").addEventListener("change", async (e) => {
    sfState.filePath = e.target.value;
    sfState.page = 1;
    await loadSfPage();
  });

  document.querySelector("#sf-page-size").addEventListener("change", async (e) => {
    sfState.pageSize = Number(e.target.value || 200);
    sfState.page = 1;
    await loadSfPage();
  });

  document.querySelector("#sf-prev").addEventListener("click", async () => {
    if (sfState.page > 1) {
      sfState.page -= 1;
      await loadSfPage();
    }
  });

  document.querySelector("#sf-next").addEventListener("click", async () => {
    sfState.page += 1;
    await loadSfPage();
  });

  document.querySelector("#sf-refresh").addEventListener("click", refreshSfAll);

  document.querySelector("#sf-search").addEventListener("input", (e) => {
    sfGridApi.setGridOption("quickFilterText", e.target.value || "");
  });

  document.querySelector("#sf-export-csv").addEventListener("click", () => {
    sfGridApi.exportDataAsCsv({ fileName: `sf_browser_page_${sfState.page}.csv` });
  });

  await refreshSfAll();
});

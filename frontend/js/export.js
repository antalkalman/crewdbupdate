let exportGridApi = null;
let projectsGridApi = null;

const exportGridOptions = {
  columnDefs: [
    {
      field: "include",
      headerName: "Include",
      width: 110,
      editable: true,
      cellRenderer: "agCheckboxCellRenderer",
      cellEditor: "agCheckboxCellEditor",
      valueParser: (params) => Boolean(params.newValue),
    },
    { field: "name", headerName: "Project Name", flex: 1, minWidth: 260 },
    { field: "id", headerName: "Project ID", width: 120 },
  ],
  rowData: [],
  defaultColDef: {
    sortable: true,
    filter: true,
    resizable: true,
  },
};

function setExportMessage(text, isError = false) {
  const el = document.querySelector("#result-message");
  el.textContent = text;
  el.style.color = isError ? "#b00020" : "#1f6d2a";
}

function selectedProjectIds() {
  const ids = [];
  exportGridApi.forEachNode((node) => {
    const row = node.data || {};
    if (row.include) {
      ids.push(Number(row.id));
    }
  });
  return ids;
}

async function loadProjectsGrid() {
  const payload = await fetchProjects();
  const included = new Set((payload.included_project_ids || []).map((v) => Number(v)));
  const rows = (payload.projects || []).map((project) => ({
    id: Number(project.id),
    name: project.name || "",
    include: included.has(Number(project.id)),
  }));
  exportGridApi.setGridOption("rowData", rows);
  setExportMessage(`Loaded ${rows.length} projects.`);
}

async function saveSelection() {
  const ids = selectedProjectIds();
  const result = await saveIncludedProjects(ids);
  setExportMessage(`Saved ${result.count} included project IDs.`);
}

async function runExport() {
  const ids = selectedProjectIds();
  const result = await runFullExport(ids);
  setExportMessage(`Export done: ${result.output_file} (${result.rows} rows)`);
}

async function runUpdateCombined() {
  const result = await updateCombinedMaster();
  const created = result.created_at || "";
  const rows = result.meta?.rows_written ?? 0;
  setExportMessage(
    `Combined master updated: ${result.output_file} at ${created} (${rows} rows written)`
  );
  document.querySelector("#go-titles").classList.remove("hidden");
  document.querySelector("#go-match").classList.remove("hidden");
}

async function runNewPipeline() {
  const result = await runNewCombine();
  const meta = result.meta || {};
  const total = meta.total_rows || 0;
  const resolved = meta.gcmid_resolved || 0;
  const mapped = meta.title_mapped || 0;
  setExportMessage(
    `New pipeline done. ` +
    `Rows: ${total.toLocaleString()}, ` +
    `GCMID resolved: ${resolved.toLocaleString()} (${total ? Math.round(resolved/total*100) : 0}%), ` +
    `Titles mapped: ${mapped.toLocaleString()} (${total ? Math.round(mapped/total*100) : 0}%). ` +
    (meta.warnings?.length ? `\u26a0 ${meta.warnings.length} warning(s).` : "")
  );
}

const projectsGridOptions = {
  columnDefs: [
    {
      field: "state",
      headerName: "State",
      width: 130,
      editable: true,
      cellEditor: "agSelectCellEditor",
      cellEditorParams: { values: ["live", "historical", "skip"] },
      cellStyle: (params) => {
        if (params.value === "live") return { color: "#1f6d2a", fontWeight: "bold" };
        if (params.value === "historical") return { color: "#555" };
        if (params.value === "skip") return { color: "#b00020" };
        return {};
      },
    },
    { field: "name", headerName: "Project", flex: 1, minWidth: 200 },
    { field: "id", headerName: "CM ID", width: 100 },
    { field: "start_date", headerName: "Start", width: 120 },
    { field: "end_date", headerName: "End", width: 120 },
  ],
  rowData: [],
  defaultColDef: { sortable: true, filter: true, resizable: true },
  rowClassRules: {
    "row-historical": (params) => params.data?.state === "historical",
    "row-skip": (params) => params.data?.state === "skip",
  },
};

async function loadProjectsManaged() {
  const data = await fetchManagedProjects();
  const stateOrder = { live: 0, historical: 1, skip: 2 };
  const sorted = (data.projects || []).slice().sort((a, b) => {
    const so = (stateOrder[a.state] ?? 9) - (stateOrder[b.state] ?? 9);
    if (so !== 0) return so;
    return (b.start_date || "").localeCompare(a.start_date || "");
  });
  projectsGridApi.setGridOption("rowData", sorted);
  const projects = data.projects || [];
  document.querySelector("#projects-message").textContent =
    `${projects.filter(p => p.state === "live").length} live, ` +
    `${projects.filter(p => p.state === "historical").length} historical, ` +
    `${projects.filter(p => p.state === "skip").length} skip`;
}

async function saveProjectStates() {
  const projects = [];
  projectsGridApi.forEachNode(node => projects.push(node.data));
  const result = await saveManagedProjects({ projects });
  document.querySelector("#projects-message").textContent =
    `Saved ${result.count} projects.`;
}

async function runLiveExportAction() {
  const result = await runLiveExport();
  setExportMessage(
    `Live export done: ${result.output_file} — ` +
    `${result.live_project_count} live projects, ${result.rows} rows`
  );
}

async function withBusy(action, message) {
  const btns = document.querySelectorAll("#run-live-export, #run-new-combine, #save-project-states");
  btns.forEach((b) => (b.disabled = true));
  setExportMessage(message);

  try {
    await action();
  } catch (error) {
    console.error(error);
    setExportMessage(error.message || "Operation failed.", true);
    alert(error.message || "Operation failed.");
  } finally {
    btns.forEach((b) => (b.disabled = false));
  }
}

document.addEventListener("DOMContentLoaded", async () => {
  const gridDiv = document.querySelector("#export-grid");
  exportGridApi = agGrid.createGrid(gridDiv, exportGridOptions);

  document.querySelector("#save-selection").addEventListener("click", () => {
    withBusy(saveSelection, "Saving selection...");
  });
  document.querySelector("#run-live-export").addEventListener("click", () => {
    withBusy(runLiveExportAction, "Exporting live projects...");
  });
  document.querySelector("#run-export").addEventListener("click", () => {
    withBusy(runExport, "Running export...");
  });
  document.querySelector("#update-combined").addEventListener("click", () => {
    withBusy(runUpdateCombined, "Updating combined master...");
  });
  document.querySelector("#run-new-combine").addEventListener("click", () => {
    withBusy(runNewPipeline, "Running new pipeline...");
  });
  document.querySelector("#refresh-projects").addEventListener("click", () => {
    withBusy(loadProjectsGrid, "Refreshing project list...");
  });

  projectsGridApi = agGrid.createGrid(
    document.querySelector("#projects-grid"),
    projectsGridOptions
  );
  document.querySelector("#save-project-states").addEventListener("click", () =>
    withBusy(saveProjectStates, "Saving project states...")
  );

  // Load managed projects first (local JSON, no API needed)
  try {
    await loadProjectsManaged();
  } catch (e) {
    console.error(e);
    document.querySelector("#projects-message").textContent = e.message || "Failed to load managed projects.";
  }

  // Load old export grid (needs API — may fail if env vars not set)
  try {
    await loadProjectsGrid();
  } catch (e) {
    console.error(e);
    setExportMessage(e.message || "Failed to load API projects.", true);
  }
});

let validGeneralTitles = [];
let gridApi = null;
let conflictsGridApi = null;
let conflictPicks = {};
let conflictsLoaded = false;

class TypeAheadGeneralTitleEditor {
  init(params) {
    this.params = params;
    this.eInput = document.createElement("input");
    this.eInput.type = "text";
    this.eInput.className = "general-title-input";
    this.eInput.style.width = "100%";
    this.eInput.value = params.value || "";

    this.listId = `general-title-options-${Math.random().toString(36).slice(2)}`;
    this.eInput.setAttribute("list", this.listId);

    this.eList = document.createElement("datalist");
    this.eList.id = this.listId;

    this.renderOptions(this.eInput.value);
    this.eInput.addEventListener("input", () => {
      this.renderOptions(this.eInput.value);
    });

    this.eContainer = document.createElement("div");
    this.eContainer.appendChild(this.eInput);
    this.eContainer.appendChild(this.eList);
  }

  renderOptions(query) {
    const q = (query || "").toLowerCase().trim();
    const options = q
      ? validGeneralTitles.filter((v) => v.toLowerCase().includes(q))
      : validGeneralTitles;
    const limited = options.slice(0, 200);
    this.eList.innerHTML = limited.map((v) => `<option value="${v}"></option>`).join("");
  }

  getGui() {
    return this.eContainer;
  }

  afterGuiAttached() {
    this.eInput.focus();
    this.eInput.select();
  }

  getValue() {
    return (this.eInput.value || "").trim();
  }
}

const gridOptions = {
  components: {
    typeAheadGeneralTitleEditor: TypeAheadGeneralTitleEditor,
  },
  columnDefs: [
    { field: "project", headerName: "Project", editable: false, flex: 1, minWidth: 120 },
    { field: "department", headerName: "Department", editable: false, flex: 1, minWidth: 120 },
    { field: "project_job_title", headerName: "Project job title", editable: false, flex: 1.5, minWidth: 150 },
    {
      field: "general_title",
      headerName: "General Title",
      editable: true,
      flex: 1.5,
      minWidth: 150,
      cellEditor: "typeAheadGeneralTitleEditor",
    },
    { field: "title_project_key", headerName: "Key", editable: false, hide: true, width: 90 },
  ],
  rowData: [],
  defaultColDef: {
    sortable: true,
    filter: true,
    resizable: true,
    flex: 0,
  },
  rowClassRules: {
    "row-mapped": (params) =>
      Boolean((params.data?.general_title || "").toString().trim()),
  },
};

function setMessage(text, isError = false) {
  const messageEl = document.querySelector("#result-message");
  messageEl.textContent = text;
  messageEl.style.color = isError ? "#b00020" : "#1f6d2a";
}

async function loadGridData() {
  const payload = await fetchUnmappedTitles();
  validGeneralTitles = Array.isArray(payload.valid_general_titles) ? payload.valid_general_titles : [];
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  gridApi.setGridOption("rowData", rows);
  setMessage(`Loaded ${rows.length} unmapped title rows.`);
}

async function submitMappings() {
  const submitButton = document.querySelector("#submit-mappings");
  const refreshButton = document.querySelector("#refresh-mappings");
  const cancelButton = document.querySelector("#cancel-mappings");
  submitButton.disabled = true;
  refreshButton.disabled = true;
  cancelButton.disabled = true;
  setMessage("Submitting mappings...");

  try {
    const rowsToSubmit = [];
    gridApi.forEachNode((node) => {
      const row = node.data || {};
      if ((row.general_title || "").toString().trim()) {
        rowsToSubmit.push(row);
      }
    });

    if (!rowsToSubmit.length) {
      setMessage("No filled General Title values to submit.");
      return;
    }

    const result = await applyTitleMappings(rowsToSubmit);
    setMessage(
      `Appended: ${result.appended}, invalid: ${result.skipped_invalid}, duplicates: ${result.skipped_duplicates}`
    );
    await loadGridData();
  } catch (error) {
    console.error(error);
    setMessage(error.message || "Submit failed.", true);
    alert(error.message || "Submit failed.");
  } finally {
    submitButton.disabled = false;
    refreshButton.disabled = false;
    cancelButton.disabled = false;
  }
}

function clearMappings() {
  let cleared = 0;
  gridApi.forEachNode((node) => {
    if ((node.data?.general_title || "").toString().trim()) {
      node.setDataValue("general_title", "");
      cleared += 1;
    }
  });
  gridApi.redrawRows();
  setMessage(`Cleared ${cleared} edited rows.`);
}

async function refreshMappings() {
  const refreshButton = document.querySelector("#refresh-mappings");
  const submitButton = document.querySelector("#submit-mappings");
  const cancelButton = document.querySelector("#cancel-mappings");
  refreshButton.disabled = true;
  submitButton.disabled = true;
  cancelButton.disabled = true;
  setMessage("Refreshing from Helper and latest SF list...");

  try {
    await loadGridData();
  } catch (error) {
    console.error(error);
    setMessage(error.message || "Refresh failed.", true);
    alert(error.message || "Refresh failed.");
  } finally {
    refreshButton.disabled = false;
    submitButton.disabled = false;
    cancelButton.disabled = false;
  }
}

// ── Tab management ──────────────────────────────────────────────────────────

function setTab(tabName) {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.tab === tabName);
  });
  document.querySelectorAll(".tab-panel").forEach((panel) => {
    panel.classList.toggle("active", panel.id === `panel-${tabName}`);
  });
}

function attachTabs() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      setTab(btn.dataset.tab);
      if (btn.dataset.tab === "conflicts" && !conflictsLoaded) {
        loadConflicts();
      }
    });
  });
}

// ── Conflicts grid ──────────────────────────────────────────────────────────

function updateWriteBtn() {
  const btn = document.querySelector("#write-conflicts-btn");
  const hasPicks = Object.values(conflictPicks).some((v) => v && v !== "skip");
  btn.disabled = !hasPicks;
}

function buildConflictsGrid() {
  const gridDiv = document.querySelector("#conflicts-grid");

  const conflictsGridOptions = {
    columnDefs: [
      {
        field: "title_project",
        headerName: "Title-Project",
        width: 280,
        pinned: "left",
        filter: true,
      },
      { field: "project", headerName: "Project", width: 140, filter: true },
      { field: "department", headerName: "Department", width: 140, filter: true },
      { field: "row_count", headerName: "Row Count", width: 90 },
      {
        field: "candidates",
        headerName: "Candidates",
        flex: 1,
        minWidth: 300,
        valueFormatter: (params) =>
          Array.isArray(params.value) ? params.value.join(" / ") : params.value || "",
        filter: true,
      },
      { field: "majority_vote", headerName: "Majority Vote", width: 200 },
      {
        headerName: "Pick",
        width: 240,
        cellRenderer: (params) => {
          const data = params.data;
          if (!data) return "";
          if (data.already_resolved) return "";

          const select = document.createElement("select");
          select.style.width = "100%";
          select.style.padding = "2px 4px";

          const defaultOpt = document.createElement("option");
          defaultOpt.value = "";
          defaultOpt.textContent = "\u2014 pick one \u2014";
          select.appendChild(defaultOpt);

          const candidates = data.candidates || [];
          candidates.forEach((c) => {
            const opt = document.createElement("option");
            opt.value = c;
            opt.textContent = c;
            select.appendChild(opt);
          });

          const skipOpt = document.createElement("option");
          skipOpt.value = "skip";
          skipOpt.textContent = "\u2715 Skip";
          select.appendChild(skipOpt);

          // Pre-select current pick or majority_vote
          const currentPick = conflictPicks[data.title_project];
          if (currentPick !== undefined) {
            select.value = currentPick;
          } else if (data.majority_vote && candidates.includes(data.majority_vote)) {
            select.value = data.majority_vote;
          }

          select.addEventListener("change", () => {
            conflictPicks[data.title_project] = select.value;
            updateWriteBtn();
          });

          return select;
        },
      },
      {
        headerName: "Status",
        width: 120,
        cellRenderer: (params) => {
          const data = params.data;
          if (!data) return "";
          if (data.already_resolved) {
            return '<span style="color:#16a34a;font-weight:600;">\u2713 Resolved</span>';
          }
          const pick = conflictPicks[data.title_project];
          if (pick && pick !== "" && pick !== "skip") {
            return '<span style="color:#2563eb;font-weight:600;">\u2713 Picked</span>';
          }
          return "";
        },
      },
    ],
    rowData: [],
    defaultColDef: {
      sortable: true,
      resizable: true,
    },
    rowClassRules: {
      "row-resolved": (params) => params.data?.already_resolved === true,
    },
    getRowStyle: (params) => {
      if (params.data?.already_resolved) {
        return { opacity: "0.5" };
      }
      return null;
    },
  };

  conflictsGridApi = agGrid.createGrid(gridDiv, conflictsGridOptions);
}

async function loadConflicts() {
  const msgEl = document.querySelector("#conflicts-message");
  msgEl.textContent = "Loading conflicts...";

  try {
    const data = await fetchTitleConflicts();
    conflictsGridApi.setGridOption("rowData", data);

    // Pre-populate picks with majority_vote
    conflictPicks = {};
    data.forEach((row) => {
      if (!row.already_resolved && row.majority_vote && row.candidates.includes(row.majority_vote)) {
        conflictPicks[row.title_project] = row.majority_vote;
      }
    });

    const unresolved = data.filter((r) => !r.already_resolved).length;
    const badge = document.querySelector("#conflicts-badge");
    badge.textContent = unresolved > 0 ? unresolved : "";

    msgEl.textContent = `${data.length} conflicts loaded, ${unresolved} unresolved.`;
    conflictsLoaded = true;
    updateWriteBtn();
  } catch (error) {
    console.error(error);
    msgEl.textContent = error.message || "Failed to load conflicts.";
  }
}

async function writeConflictPicks() {
  const btn = document.querySelector("#write-conflicts-btn");
  const msgEl = document.querySelector("#conflicts-message");
  btn.disabled = true;
  msgEl.textContent = "Writing resolved mappings...";

  try {
    const rowsToWrite = [];
    for (const [key, pick] of Object.entries(conflictPicks)) {
      if (pick && pick !== "" && pick !== "skip") {
        rowsToWrite.push({
          title_project_key: key,
          general_title: pick,
        });
      }
    }

    if (!rowsToWrite.length) {
      msgEl.textContent = "No picks to write.";
      btn.disabled = false;
      return;
    }

    const result = await applyTitleMappings(rowsToWrite);
    msgEl.textContent = `${result.appended} mappings written to Helper.xlsx` +
      (result.skipped_duplicates ? `, ${result.skipped_duplicates} already existed` : "");

    // Refresh to update status column
    await loadConflicts();
  } catch (error) {
    console.error(error);
    msgEl.textContent = error.message || "Write failed.";
  } finally {
    btn.disabled = false;
    updateWriteBtn();
  }
}

// ── Init ────────────────────────────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", async () => {
  const gridDiv = document.querySelector("#title-grid");
  gridApi = agGrid.createGrid(gridDiv, gridOptions);

  buildConflictsGrid();
  attachTabs();

  document.querySelector("#submit-mappings").addEventListener("click", submitMappings);
  document.querySelector("#refresh-mappings").addEventListener("click", refreshMappings);
  document.querySelector("#cancel-mappings").addEventListener("click", clearMappings);
  document.querySelector("#write-conflicts-btn").addEventListener("click", writeConflictPicks);

  try {
    await loadGridData();
  } catch (error) {
    console.error(error);
    setMessage(error.message || "Failed to load data.", true);
    alert(error.message || "Failed to load data.");
  }
});

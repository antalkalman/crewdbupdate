let gridApi = null;
let allRows = [];
let allColumns = [];
let extraCols = [];
let pendingChanges = {};
let saveTimer = null;

const FIXED_COLS = [
  "Sf number", "Crew member id", "State", "Project",
  "Project department", "Project job title", "Project unit",
  "Crew list name", "Surname", "Firstname",
  "Mobile number", "Crew email",
  "Start date", "End date", "Deal type",
  "Project overtime", "Project turnaround", "Project working hour",
  "Daily fee", "Weekly fee", "Note",
];

function setMessage(text, isError = false) {
  const el = document.querySelector("#sf-issues-message");
  el.textContent = text;
  el.style.color = isError ? "#b00020" : "#1f6d2a";
}

function buildColumnDefs() {
  const cols = [];

  // Checkbox column
  cols.push({
    headerName: "",
    field: "checked",
    width: 48,
    pinned: "left",
    checkboxSelection: true,
    headerCheckboxSelection: true,
    sortable: false,
    filter: false,
    resizable: false,
  });

  // Fixed columns
  FIXED_COLS.forEach((col) => {
    const def = {
      field: col,
      headerName: col,
      sortable: true,
      filter: true,
      resizable: true,
    };
    if (col === "Note") {
      def.editable = true;
      def.width = 400;
      def.cellStyle = (p) => (p.value ? { color: "#b00020", fontStyle: "italic" } : {});
      def.wrapText = true;
      def.autoHeight = true;
    } else if (col === "Sf number") {
      def.width = 110;
      def.pinned = "left";
    } else if (col === "Crew member id") {
      def.width = 100;
    } else if (col === "State") {
      def.width = 90;
      def.cellStyle = (p) => {
        const v = (p.value || "").toLowerCase();
        if (v === "signed") return { color: "#1f6d2a" };
        if (v === "accepted") return { color: "#1a5fa8" };
        return { color: "#b00020" };
      };
    } else if (["Start date", "End date"].includes(col)) {
      def.width = 110;
    } else if (["Daily fee", "Weekly fee"].includes(col)) {
      def.width = 100;
    } else {
      def.width = 160;
    }
    cols.push(def);
  });

  // Extra columns added by user
  extraCols.forEach((col) => {
    cols.push({
      field: col,
      headerName: col,
      sortable: true,
      filter: true,
      resizable: true,
      width: 160,
    });
  });

  return cols;
}

function onNoteChanged(params) {
  const sfNum = params.data["Sf number"];
  if (!pendingChanges[sfNum]) pendingChanges[sfNum] = {};
  pendingChanges[sfNum].sf_number = sfNum;
  pendingChanges[sfNum].checked = params.data.checked || false;
  pendingChanges[sfNum].note = params.newValue || "";
  pendingChanges[sfNum].note_edited = true;
  params.data._note_edited = true;
  scheduleSave();
}

function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    const changes = Object.values(pendingChanges);
    if (changes.length) {
      try {
        await saveSfIssuesState(changes);
        pendingChanges = {};
      } catch (_) {}
    }
  }, 1500);
}

function updateExportButtons() {
  const checkedRows = gridApi.getSelectedRows();
  const issueCount = allRows.filter((r) => r.has_issues).length;
  document.querySelector("#export-selected").disabled = checkedRows.length === 0;
  document.querySelector("#export-all-issues").disabled = issueCount === 0;
}

function applyFilters() {
  const project = document.querySelector("#project-filter").value;
  const issuesOnly = document.querySelector("#issues-only").checked;

  let filtered = allRows;
  if (project) filtered = filtered.filter((r) => r.Project === project);
  if (issuesOnly) filtered = filtered.filter((r) => r.has_issues);

  gridApi.setGridOption("rowData", filtered);
  document.querySelector("#sf-issues-info").textContent =
    `Showing ${filtered.length} of ${allRows.length} SFs \u00b7 ` +
    `${allRows.filter((r) => r.has_issues).length} with issues \u00b7 ` +
    `${gridApi.getSelectedRows().length} selected`;
}

function populateProjectFilter(projects) {
  const sel = document.querySelector("#project-filter");
  sel.innerHTML = '<option value="">All projects</option>';
  projects.forEach((p) => {
    const opt = document.createElement("option");
    opt.value = p;
    opt.textContent = p;
    sel.appendChild(opt);
  });
}

// ── Extra column picker ──────────────────────────────────────────────────────

function renderExtraColTags() {
  const container = document.querySelector("#extra-col-tags");
  container.innerHTML = "";
  extraCols.forEach((col) => {
    const tag = document.createElement("span");
    tag.className = "filter-tag";
    const label = document.createTextNode(col + " ");
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = "\u00d7";
    btn.addEventListener("click", () => {
      extraCols = extraCols.filter((c) => c !== col);
      renderExtraColTags();
      gridApi.setGridOption("columnDefs", buildColumnDefs());
    });
    tag.appendChild(label);
    tag.appendChild(btn);
    container.appendChild(tag);
  });
}

function setupExtraColPicker(availableCols) {
  const input = document.querySelector("#extra-col-picker");
  const dropdown = document.querySelector("#extra-col-dropdown");
  const addable = availableCols.filter((c) => !FIXED_COLS.includes(c) && c !== "Note");

  function renderDropdown(query) {
    const filtered = addable
      .filter((c) => !extraCols.includes(c))
      .filter((c) => c.toLowerCase().includes(query.toLowerCase()));
    if (!filtered.length) {
      dropdown.style.display = "none";
      return;
    }
    dropdown.innerHTML = "";
    filtered.slice(0, 40).forEach((col) => {
      const item = document.createElement("div");
      item.className = "filter-dropdown-item";
      item.textContent = col;
      item.addEventListener("mousedown", (e) => {
        e.preventDefault();
        extraCols.push(col);
        renderExtraColTags();
        gridApi.setGridOption("columnDefs", buildColumnDefs());
        input.value = "";
        dropdown.style.display = "none";
      });
      dropdown.appendChild(item);
    });
    dropdown.style.display = "block";
  }

  input.addEventListener("input", (e) => renderDropdown(e.target.value));
  input.addEventListener("focus", (e) => renderDropdown(e.target.value));
  input.addEventListener("blur", () => setTimeout(() => (dropdown.style.display = "none"), 150));
}

// ── Init ─────────────────────────────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", () => {
  gridApi = agGrid.createGrid(document.querySelector("#sf-issues-grid"), {
    columnDefs: [],
    rowData: [],
    defaultColDef: { sortable: true, filter: true, resizable: true },
    rowSelection: "multiple",
    suppressRowClickSelection: true,
    onSelectionChanged: () => updateExportButtons(),
    onCellValueChanged: (params) => {
      if (params.column.getColId() === "Note") {
        onNoteChanged(params);
      }
    },
    rowClassRules: {
      "row-issues": (p) => p.data?.has_issues,
    },
  });

  document.querySelector("#run-check").addEventListener("click", async () => {
    setMessage("Running checks...");
    document.querySelector("#run-check").disabled = true;
    try {
      const result = await runSfIssuesCheck();
      allRows = result.rows;
      allColumns = result.all_columns || result.columns;

      setupExtraColPicker(allColumns);
      gridApi.setGridOption("columnDefs", buildColumnDefs());
      applyFilters();
      populateProjectFilter(result.projects);
      updateExportButtons();
      setMessage(`${result.source_file} \u00b7 ${result.total} SFs \u00b7 ${result.with_issues} with issues`);
    } catch (err) {
      setMessage(err.message || "Check failed", true);
    } finally {
      document.querySelector("#run-check").disabled = false;
    }
  });

  document.querySelector("#project-filter").addEventListener("change", applyFilters);
  document.querySelector("#issues-only").addEventListener("change", applyFilters);

  document.querySelector("#export-selected").addEventListener("click", async () => {
    const selected = gridApi.getSelectedRows();
    if (!selected.length) return;
    const today = new Date().toISOString().slice(0, 10);
    try {
      await exportSfIssues(selected, `SF_Issues_Selected_${today}`);
      setMessage("Export downloaded.");
    } catch (err) {
      setMessage(err.message || "Export failed", true);
    }
  });

  document.querySelector("#export-all-issues").addEventListener("click", async () => {
    const issues = allRows.filter((r) => r.has_issues);
    if (!issues.length) return;
    const today = new Date().toISOString().slice(0, 10);
    try {
      await exportSfIssues(issues, `SF_Issues_All_${today}`);
      setMessage("Export downloaded.");
    } catch (err) {
      setMessage(err.message || "Export failed", true);
    }
  });
});

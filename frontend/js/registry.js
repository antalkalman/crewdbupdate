let registryGridApi = null;
let pendingChanges = {};   // keyed by "cm_id|field" -> {cm_id, field, value}
let generalTitles = [];

const EDITABLE_COLS = ["Sure Name", "First Name", "Nick Name", "Actual Title", "Status", "Actual Phone", "Actual Email", "Note"];

function formatPhone(val) {
  if (!val) return "";
  const s = String(val).replace(/\D/g, "");
  if (s.length < 8) return s;
  // +CC AA BBBB rest  (e.g. +36 30 5290 050)
  const cc = s.slice(0, 2);
  const area = s.slice(2, 4);
  const block = s.slice(4, 8);
  const rest = s.slice(8);
  return `+${cc} ${area} ${block} ${rest}`.trim();
}
const AUTO_COLS = ["Last General Title", "Last Department", "Title Flag", "Last Email", "Last Phone", "Shows Worked", "Actual Name"];
const STATUS_VALUES = ["Active", "Retired", "Foreign", "External"];

function buildColumnDefs(headers) {
  return headers.map((col) => {
    const isEditable = EDITABLE_COLS.includes(col);
    const isAuto = AUTO_COLS.includes(col);
    const isId = col === "CM ID";

    const def = {
      field: col,
      headerName: col,
      editable: isEditable,
      resizable: true,
      sortable: true,
      filter: true,
    };

    // Column widths
    if (isId) def.width = 80;
    else if (col === "Shows Worked") def.width = 300;
    else if (["Sure Name", "First Name", "Actual Name"].includes(col)) def.width = 160;
    else if (["Actual Title", "Last General Title", "Last Department"].includes(col)) def.width = 200;
    else if (col === "Note") def.width = 200;
    else if (col === "Status") def.width = 110;
    else if (col === "Actual Phone") {
      def.width = 160;
      def.valueFormatter = (p) => formatPhone(p.value);
    }
    else if (col === "Actual Email") def.width = 220;
    else if (col === "Last Phone") {
      def.width = 160;
      def.valueFormatter = (p) => formatPhone(p.value);
    }
    else def.width = 130;

    // Auto cols — grey background, not editable
    if (isAuto) {
      def.cellStyle = { background: "#f5f5f5", color: "#666" };
    }

    // Status — dropdown editor
    if (col === "Status") {
      def.cellEditor = "agSelectCellEditor";
      def.cellEditorParams = { values: STATUS_VALUES };
      def.cellStyle = (params) => {
        const v = params.value;
        if (v === "Active") return { color: "#1f6d2a", fontWeight: "600" };
        if (v === "Retired") return { color: "#888" };
        if (v === "Foreign") return { color: "#1a5fa8" };
        if (v === "External") return { color: "#b00020" };
        return {};
      };
    }

    // Actual Title — dropdown editor with General Titles
    if (col === "Actual Title") {
      def.cellEditor = "agSelectCellEditor";
      def.cellEditorParams = () => ({ values: ["", ...generalTitles] });
    }

    return def;
  });
}

function onCellValueChanged(params) {
  const cmId = params.data["CM ID"];
  const field = params.column.getColId();
  const value = params.newValue ?? "";
  const key = `${cmId}|${field}`;
  pendingChanges[key] = { cm_id: Number(cmId), field, value: String(value) };
  updatePendingUI();
}

function updatePendingUI() {
  const count = Object.keys(pendingChanges).length;
  const saveBtn = document.querySelector("#save-changes");
  const pendingEl = document.querySelector("#pending-count");
  saveBtn.disabled = count === 0;
  pendingEl.textContent = count > 0
    ? `${count} unsaved change${count > 1 ? "s" : ""}`
    : "";
}

async function loadRegistry() {
  const msgEl = document.querySelector("#registry-message");
  msgEl.textContent = "Loading...";
  msgEl.style.color = "";

  try {
    // Load general titles for dropdown
    if (generalTitles.length === 0) {
      try {
        const gt = await fetchGeneralTitles();
        generalTitles = gt.titles || [];
      } catch (_) {
        generalTitles = [];
      }
    }

    const data = await fetchRegistry();
    const colDefs = buildColumnDefs(data.columns);

    registryGridApi.updateGridOptions({ columnDefs: colDefs });
    registryGridApi.setGridOption("rowData", data.rows);

    // Default filter: Status = Active only
    registryGridApi.setFilterModel({
      Status: { filterType: "text", type: "equals", filter: "Active" },
    });

    pendingChanges = {};
    updatePendingUI();
    msgEl.textContent = `Loaded ${data.rows.length.toLocaleString()} crew members.`;
  } catch (err) {
    msgEl.style.color = "#b00020";
    msgEl.textContent = err.message || "Failed to load registry.";
  }
}

async function saveChanges() {
  const changes = Object.values(pendingChanges);
  if (!changes.length) return;

  const msgEl = document.querySelector("#registry-message");
  msgEl.style.color = "";
  msgEl.textContent = `Saving ${changes.length} changes...`;

  try {
    const result = await saveRegistry(changes);
    pendingChanges = {};
    updatePendingUI();
    msgEl.style.color = "#1f6d2a";
    msgEl.textContent = `Saved ${result.saved} changes to CrewRegistry.xlsx`;
  } catch (err) {
    msgEl.style.color = "#b00020";
    msgEl.textContent = err.message || "Failed to save changes.";
  }
}

document.addEventListener("DOMContentLoaded", async () => {
  registryGridApi = agGrid.createGrid(
    document.querySelector("#registry-grid"),
    {
      columnDefs: [],
      rowData: [],
      defaultColDef: { resizable: true, sortable: true, filter: true },
      onCellValueChanged,
      rowClassRules: {
        "row-retired": (p) => p.data?.Status === "Retired",
        "row-external": (p) => p.data?.Status === "External",
        "row-foreign": (p) => p.data?.Status === "Foreign",
      },
    }
  );

  document.querySelector("#save-changes").addEventListener("click", saveChanges);
  document.querySelector("#refresh-registry").addEventListener("click", () => {
    pendingChanges = {};
    loadRegistry();
  });

  await loadRegistry();
});

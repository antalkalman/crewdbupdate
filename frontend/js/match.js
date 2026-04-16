// ── Module-level state ────────────────────────────────────────────────────────
let confirmedGridApi = null;
let possibleGridApi  = null;
let newNamesGridApi  = null;

const expandedRows   = new Set();   // source_key values currently expanded
let   possibleBaseRows = [];        // original possible rows from last match run
const possiblePicks  = {};          // source_key → { cm_job, cm_id } | "none"

// ── New Name Modal state ─────────────────────────────────────────────────────
let modalRows = [];       // edited copies of selected rows
let modalIndex = 0;       // current position in the stepper
let generalTitles = [];   // loaded once from GET /api/general_titles

// ── Utilities ─────────────────────────────────────────────────────────────────
function setMessage(text, isError = false) {
  const el = document.querySelector("#result-message");
  el.textContent = text;
  el.style.color = isError ? "#b00020" : "#1f6d2a";
}

// ── Possible: Full-Width Detail Renderer ──────────────────────────────────────
class PossibleDetailRenderer {
  init(params) {
    const parentKey  = params.data.__parentKey;
    const candidates = params.data.__candidates || [];

    this.eGui = document.createElement("div");
    this.eGui.style.cssText =
      "padding:8px 16px 12px 40px; background:#f8fafc; border-bottom:2px solid #cbd5e1;";

    const table = document.createElement("table");
    table.style.cssText =
      "width:100%; border-collapse:collapse; font-size:12px;";

    // ── Header ────────────────────────────────────────────────────────────────
    const thead = document.createElement("thead");
    thead.innerHTML = `<tr style="background:#e2e8f0; text-align:left;">
      <th style="padding:5px 8px;">#</th>
      <th style="padding:5px 8px;">GCMID</th>
      <th style="padding:5px 8px;">Name</th>
      <th style="padding:5px 8px;">Title</th>
      <th style="padding:5px 8px;">Dept</th>
      <th style="padding:5px 8px; text-align:right;">Name</th>
      <th style="padding:5px 8px; text-align:right;">Email</th>
      <th style="padding:5px 8px; text-align:right;">Phone</th>
      <th style="padding:5px 8px; text-align:right;">Dept</th>
      <th style="padding:5px 8px; text-align:right;">Final</th>
      <th style="padding:5px 8px;"></th>
    </tr>`;
    table.appendChild(thead);

    // ── Body ──────────────────────────────────────────────────────────────────
    const tbody = document.createElement("tbody");
    this._candidateRows = [];

    candidates.forEach((c, i) => {
      const tr = document.createElement("tr");
      tr.style.background = i % 2 === 0 ? "#ffffff" : "#f1f5f9";
      tr.style.transition  = "background 0.15s";
      this._candidateRows.push(tr);

      // Check if already picked
      const currentPick = possiblePicks[parentKey];
      if (currentPick && currentPick !== "none" &&
          currentPick.cm_id === c.suggested_gcmid) {
        tr.style.background  = "#dcfce7";
        tr.style.fontWeight  = "bold";
      }

      const f = (v) => (v == null ? "" : Number(v).toFixed(2));
      const s = (v) => (v == null ? "" : String(v));

      tr.innerHTML = `
        <td style="padding:5px 8px;">${c.rank ?? i + 1}</td>
        <td style="padding:5px 8px;">${s(c.suggested_gcmid)}</td>
        <td style="padding:5px 8px; white-space:nowrap;">${s(c.matched_name)}</td>
        <td style="padding:5px 8px;">${s(c.db_title)}</td>
        <td style="padding:5px 8px;">${s(c.db_general_department)}</td>
        <td style="padding:5px 8px; text-align:right;">${f(c.name_score)}</td>
        <td style="padding:5px 8px; text-align:right;">${f(c.email_score)}</td>
        <td style="padding:5px 8px; text-align:right;">${f(c.phone_score)}</td>
        <td style="padding:5px 8px; text-align:right;">${f(c.dept_score)}</td>
        <td style="padding:5px 8px; text-align:right;">${f(c.final_score)}</td>
        <td style="padding:5px 8px;"></td>`;

      const pickBtn = document.createElement("button");
      pickBtn.type      = "button";
      pickBtn.textContent = "✓ Pick";
      pickBtn.style.cssText =
        "background:#16a34a;color:white;border:none;padding:3px 10px;" +
        "border-radius:4px;cursor:pointer;font-size:12px;";

      pickBtn.addEventListener("click", () => {
        // Store pick
        possiblePicks[parentKey] = { cm_job: parentKey, cm_id: c.suggested_gcmid };

        // Highlight selected row, reset others
        this._candidateRows.forEach((r, idx) => {
          if (r === tr) {
            r.style.background = "#dcfce7";
            r.style.fontWeight = "bold";
          } else {
            r.style.background = idx % 2 === 0 ? "#ffffff" : "#f1f5f9";
            r.style.fontWeight = "";
          }
        });

        // Enable write button if any real pick exists
        const hasPicks = Object.values(possiblePicks).some(
          (v) => v !== "none" && v?.cm_job
        );
        document.querySelector("#write-possible-btn").disabled = !hasPicks;

        // Refresh parent row status column
        possibleGridApi.refreshCells({ force: true });
      });

      tr.querySelector("td:last-child").appendChild(pickBtn);
      tbody.appendChild(tr);
    });

    // ── "None of these" + "Add as New" row ─────────────────────────────────────
    const noneRow = document.createElement("tr");
    noneRow.style.borderTop = "1px solid #cbd5e1";
    const noneTd = document.createElement("td");
    noneTd.colSpan = 11;
    noneTd.style.padding = "8px";

    const noneBtn = document.createElement("button");
    noneBtn.type        = "button";
    noneBtn.textContent = "✕ None of these";
    noneBtn.style.cssText =
      "padding:5px 12px;background:#f5f5f5;border:1px solid #ccc;" +
      "border-radius:4px;cursor:pointer;font-size:12px;margin-right:8px;";

    noneBtn.addEventListener("click", () => {
      possiblePicks[parentKey] = "none";
      possibleGridApi.refreshCells({ force: true });
    });

    const addNewBtn = document.createElement("button");
    addNewBtn.type        = "button";
    addNewBtn.textContent = "＋ Add as New";
    addNewBtn.style.cssText =
      "padding:5px 12px;background:#1a6b3a;color:#fff;border:none;" +
      "border-radius:4px;cursor:pointer;font-size:12px;";

    addNewBtn.addEventListener("click", () => {
      const syntheticRow = {
        source_key:        parentKey,
        name_on_crew_list: params.data.__parentName || "",
        project_job_title: params.data.__parentJobTitle || "",
        phone:             params.data.__parentPhone || "",
        email:             params.data.__parentEmail || "",
        general_department: params.data.__parentDept || "",
        general_title:     params.data.__parentGeneralTitle || "",
      };
      openNewNameModal([syntheticRow]);
    });

    noneTd.appendChild(noneBtn);
    noneTd.appendChild(addNewBtn);
    noneRow.appendChild(noneTd);
    tbody.appendChild(noneRow);

    table.appendChild(tbody);
    this.eGui.appendChild(table);
  }

  getGui() {
    return this.eGui;
  }
}

// ── Possible: row array builder ───────────────────────────────────────────────
function rebuildPossibleRows() {
  const rows = [];
  for (const row of possibleBaseRows) {
    rows.push(row);
    if (expandedRows.has(row.source_key)) {
      rows.push({
        __isDetail:    true,
        __parentKey:   row.source_key,
        __parentName:  row.name_on_crew_list,
        __parentJobTitle: row.project_job_title,
        __parentPhone: row.phone,
        __parentEmail: row.email,
        __parentDept:  row.general_department,
        __parentGeneralTitle: row.general_title,
        __candidates:  row.candidates || [],
      });
    }
  }
  return rows;
}

// ── Grid builders ─────────────────────────────────────────────────────────────
function buildConfirmedGrid() {
  return {
    columnDefs: [
      { field: "source_key", headerName: "Source Key", width: 190,
        checkboxSelection: true, headerCheckboxSelection: true },
      { field: "name_on_crew_list", headerName: "Name",      width: 220 },
      { field: "project",           headerName: "Project",   width: 150 },
      { field: "project_job_title", headerName: "Job Title", width: 190 },
      { field: "suggested_gcmid",   headerName: "Suggested GCMID", width: 150 },
      { field: "matched_name",      headerName: "Matched Name",    width: 200 },
      { field: "db_title",          headerName: "DB Title",        width: 180 },
      { field: "name_score",        headerName: "Name",   width: 90 },
      { field: "email_score",       headerName: "Email",  width: 90 },
      { field: "phone_score",       headerName: "Phone",  width: 90 },
      { field: "dept_score",        headerName: "Dept",   width: 90 },
      { field: "final_score",       headerName: "Final",  width: 90 },
    ],
    rowData: [],
    defaultColDef: { sortable: true, filter: true, resizable: true },
    rowSelection: "multiple",
    suppressRowClickSelection: true,
    onSelectionChanged: () => {
      const count = confirmedGridApi.getSelectedRows().length;
      document.querySelector("#confirm-selected-btn").disabled = count === 0;
    },
  };
}

function buildPossibleGrid() {
  return {
    columnDefs: [
      {
        headerName: "",
        width: 40,
        sortable: false,
        filter: false,
        resizable: false,
        cellRenderer: (params) => {
          if (params.data?.__isDetail) return "";
          const key = params.data?.source_key;
          const btn = document.createElement("button");
          btn.type = "button";
          btn.textContent = expandedRows.has(key) ? "▼" : "▶";
          btn.style.cssText =
            "background:none;border:none;cursor:pointer;font-size:13px;padding:0 2px;";
          btn.addEventListener("click", (e) => {
            e.stopPropagation();
            if (expandedRows.has(key)) {
              expandedRows.delete(key);
            } else {
              expandedRows.add(key);
            }
            possibleGridApi.setGridOption("rowData", rebuildPossibleRows());
          });
          return btn;
        },
      },
      {
        headerName: "Status",
        width: 90,
        sortable: false,
        filter: false,
        resizable: false,
        cellRenderer: (params) => {
          if (params.data?.__isDetail) return "";
          const pick = possiblePicks[params.data?.source_key];
          if (pick === "added") {
            const span = document.createElement("span");
            span.textContent   = "✓ Added";
            span.style.color   = "#16a34a";
            span.style.fontWeight = "bold";
            return span;
          }
          if (pick && pick !== "none") {
            const span = document.createElement("span");
            span.textContent   = "✓ Picked";
            span.style.color   = "#16a34a";
            span.style.fontWeight = "bold";
            return span;
          }
          if (pick === "none") {
            const span = document.createElement("span");
            span.textContent = "✕ Skipped";
            span.style.color = "#94a3b8";
            return span;
          }
          return "";
        },
      },
      { field: "source_key",        headerName: "Source Key", width: 200 },
      { field: "name_on_crew_list", headerName: "Name",       width: 220 },
      { field: "project",           headerName: "Project",    width: 150 },
      { field: "project_job_title", headerName: "Job Title",  width: 190 },
      { field: "candidate_count",   headerName: "Candidates", width: 100 },
      {
        headerName: "Best Score ↓",
        width: 105,
        sortable: false,
        filter: false,
        resizable: true,
        valueGetter: (params) => {
          if (params.data?.__isDetail) return null;
          const candidates = params.data?.candidates || [];
          if (!candidates.length) return null;
          return Math.max(...candidates.map((c) => c.final_score ?? 0));
        },
        valueFormatter: (params) =>
          params.value != null ? params.value.toFixed(2) : "",
      },
      { field: "phone", headerName: "Phone", width: 150,
        valueFormatter: (params) => {
          if (params.data?.__isDetail || !params.value) return "";
          const s = String(params.value).replace(/\D/g, "");
          if (s.length < 8) return s;
          return `+${s.slice(0,2)} ${s.slice(2,4)} ${s.slice(4,8)} ${s.slice(8)}`.trim();
        },
      },
      { field: "email", headerName: "Email", width: 200 },
    ],
    rowData: [],
    defaultColDef: { sortable: false, filter: true, resizable: true },
    suppressRowClickSelection: true,
    isFullWidthRow:       (params) => params.rowNode.data?.__isDetail === true,
    fullWidthCellRenderer: PossibleDetailRenderer,
    getRowHeight: (params) => {
      if (params.node.data?.__isDetail) {
        const count = (params.node.data.__candidates || []).length;
        return count * 42 + 130;
      }
      return 48;
    },
  };
}

function buildNewNamesGrid() {
  return {
    columnDefs: [
      { field: "source_key",         headerName: "Source Key", width: 190,
        checkboxSelection: true, headerCheckboxSelection: true },
      { field: "name_on_crew_list",  headerName: "Name",       width: 220 },
      { field: "project",            headerName: "Project",    width: 150 },
      { field: "project_job_title",  headerName: "Job Title",  width: 190 },
      { field: "general_department", headerName: "Dept",       width: 140 },
      { field: "general_title",      headerName: "Title",      width: 170 },
      { field: "phone",              headerName: "Phone",      width: 150 },
      { field: "email",              headerName: "Email",      width: 220 },
      { field: "reason",             headerName: "Reason",     flex: 1, minWidth: 260 },
    ],
    rowData: [],
    defaultColDef: { sortable: true, filter: true, resizable: true },
    rowSelection: "multiple",
    suppressRowClickSelection: true,
    onSelectionChanged: () => {
      const count = newNamesGridApi.getSelectedRows().length;
      document.querySelector("#add-new-names-btn").disabled = count === 0;
    },
  };
}

// ── Tab management ────────────────────────────────────────────────────────────
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
    btn.addEventListener("click", () => setTab(btn.dataset.tab));
  });
}

// ── Actions ───────────────────────────────────────────────────────────────────
async function confirmSelected() {
  const rows = confirmedGridApi.getSelectedRows();
  if (!rows.length) return;
  const entries = rows.map((r) => ({
    cm_job: r.source_key,
    cm_id:  Number(r.suggested_gcmid),
  }));
  const msgEl = document.querySelector("#confirm-message");
  msgEl.style.color = "";
  msgEl.textContent = `Writing ${entries.length} entries...`;
  try {
    const result = await confirmMatchesToGcmidMap(entries);
    msgEl.style.color   = "#1f6d2a";
    msgEl.textContent   = `Done: ${result.added} added, ${result.skipped} skipped (already existed).`;
  } catch (err) {
    msgEl.style.color   = "#b00020";
    msgEl.textContent   = err.message || "Failed to write to GCMID Map.";
  }
}

async function writePossiblePicks() {
  const entries = Object.values(possiblePicks)
    .filter((v) => v !== "none" && v?.cm_job)
    .map((v) => ({ cm_job: v.cm_job, cm_id: Number(v.cm_id) }));
  if (!entries.length) return;
  const msgEl = document.querySelector("#possible-message");
  msgEl.textContent   = `Writing ${entries.length} picks...`;
  msgEl.style.color   = "";
  try {
    const result = await confirmMatchesToGcmidMap(entries);
    msgEl.style.color   = "#1f6d2a";
    msgEl.textContent   = `Done: ${result.added} added, ${result.skipped} skipped.`;
  } catch (err) {
    msgEl.style.color   = "#b00020";
    msgEl.textContent   = err.message || "Failed.";
  }
}

// ── New Name Review Modal ─────────────────────────────────────────────────────

async function openNewNameModal(selectedRows) {
  // Load general titles once per session
  if (generalTitles.length === 0) {
    try {
      const gt = await fetchGeneralTitles();
      generalTitles = gt.titles || [];
    } catch (e) {
      generalTitles = [];
    }
  }

  // Build editable copies with name pre-split (last word = surname)
  const BLOCKED_EMAILS = ["pioneer@crewcall.hu"];
  modalRows = selectedRows.map((r) => {
    const name = (r.name_on_crew_list || "").trim();
    const words = name.split(" ");
    const surname = words.length > 1 ? words[words.length - 1] : "";
    const firstname = words.length > 1 ? words.slice(0, -1).join(" ") : name;
    return {
      source_key: r.source_key,
      firstname,
      surname,
      raw_title: r.project_job_title || "",
      actual_title: r.general_title || "",
      phone: (r.phone && r.phone !== "nan") ? r.phone : "",
      email: (r.email && r.email !== "nan" && !BLOCKED_EMAILS.includes((r.email || "").toLowerCase())) ? r.email : "",
      general_department: r.general_department || "",
      status: "Active",
      note: "",
    };
  });

  modalIndex = 0;

  // Populate title dropdown
  const titleSelect = document.querySelector("#modal-title");
  titleSelect.innerHTML = '<option value="">-- leave blank --</option>';
  generalTitles.forEach((t) => {
    const opt = document.createElement("option");
    opt.value = t;
    opt.textContent = t;
    titleSelect.appendChild(opt);
  });

  renderModalRow();
  document.querySelector("#new-name-modal").style.display = "flex";
}

function renderModalRow() {
  const row = modalRows[modalIndex];
  const total = modalRows.length;

  document.querySelector("#modal-counter").textContent = `${modalIndex + 1} of ${total}`;
  document.querySelector("#modal-source-key").textContent = row.source_key;
  document.querySelector("#modal-firstname").value = row.firstname;
  document.querySelector("#modal-surname").value = row.surname;
  document.querySelector("#modal-title").value = row.actual_title;
  document.querySelector("#modal-raw-title").textContent = row.raw_title;
  document.querySelector("#modal-phone").value = row.phone;
  document.querySelector("#modal-email").value = row.email;
  document.querySelector("#modal-status").value = row.status;
  document.querySelector("#modal-note").value = row.note;

  const nextBtn = document.querySelector("#modal-next");
  nextBtn.textContent = modalIndex === total - 1 ? "Confirm & Write All" : "Next \u2192";

  document.querySelector("#modal-back").style.visibility =
    modalIndex === 0 ? "hidden" : "visible";
}

function saveCurrentModalRow() {
  const row = modalRows[modalIndex];
  row.firstname = document.querySelector("#modal-firstname").value.trim();
  row.surname = document.querySelector("#modal-surname").value.trim();
  row.actual_title = document.querySelector("#modal-title").value;
  row.phone = document.querySelector("#modal-phone").value.trim();
  row.email = document.querySelector("#modal-email").value.trim();
  row.status = document.querySelector("#modal-status").value;
  row.note = document.querySelector("#modal-note").value.trim();
}

async function submitModalRows() {
  const entries = modalRows.map((r) => ({
    source_key: r.source_key,
    name_on_crew_list: `${r.firstname} ${r.surname}`.trim(),
    project_job_title: r.raw_title,
    actual_title_override: r.actual_title,
    phone: r.phone,
    email: r.email,
    general_department: r.general_department,
    status: r.status,
    note: r.note,
  }));

  const msgEl = document.querySelector("#new-names-message");
  msgEl.style.color = "";
  msgEl.textContent = `Adding ${entries.length} new records...`;
  try {
    const result = await addToRegistry(entries);
    document.querySelector("#new-name-modal").style.display = "none";

    // Mark source keys as "added" in possiblePicks so Possible tab shows "✓ Added"
    modalRows.forEach((r) => {
      if (possiblePicks.hasOwnProperty(r.source_key) || possibleBaseRows.some((p) => p.source_key === r.source_key)) {
        possiblePicks[r.source_key] = "added";
      }
    });
    if (possibleGridApi) possibleGridApi.refreshCells({ force: true });

    msgEl.style.color = "#1f6d2a";
    msgEl.textContent = `Done: ${result.added} added. New CM IDs: ${result.new_ids.join(", ")}`;
  } catch (err) {
    msgEl.style.color = "#b00020";
    msgEl.textContent = err.message || "Failed to write to CrewRegistry";
  }
}

async function addSelectedNewNames() {
  const rows = newNamesGridApi.getSelectedRows();
  if (!rows.length) return;
  await openNewNameModal(rows);
}

async function runMatching() {
  const runBtn = document.querySelector("#run-matching");
  runBtn.disabled = true;
  setMessage("Running matching...");

  try {
    const result = await runMatchJob();

    confirmedGridApi.setGridOption("rowData", result.confirmed || []);
    newNamesGridApi.setGridOption("rowData",  result.new_names  || []);
    document.querySelector("#add-new-names-btn").disabled  = true;
    document.querySelector("#new-names-message").textContent = "";

    // Reset Possible state
    Object.keys(possiblePicks).forEach((k) => delete possiblePicks[k]);
    expandedRows.clear();
    possibleBaseRows = (result.possible || []).slice().sort((a, b) => {
      const bestScore = (row) => {
        const c = row.candidates || [];
        return c.length ? Math.max(...c.map((x) => x.final_score ?? 0)) : -Infinity;
      };
      return bestScore(b) - bestScore(a);
    });
    possibleGridApi.setGridOption("rowData", possibleBaseRows);
    document.querySelector("#write-possible-btn").disabled  = true;
    document.querySelector("#possible-message").textContent = "";

    const meta = result.meta || {};
    setMessage(
      `Done. Missing: ${meta.missing_count || 0}, confirmed: ${meta.confirmed_count || 0}, ` +
      `possible: ${meta.possible_count || 0}, new: ${meta.new_names_count || 0}`
    );
  } catch (error) {
    console.error(error);
    setMessage(error.message || "Match run failed.", true);
    alert(error.message || "Match run failed.");
  } finally {
    runBtn.disabled = false;
  }
}

async function runNewMatching() {
  const runBtn = document.querySelector("#run-matching");
  const newRunBtn = document.querySelector("#run-new-matching");
  runBtn.disabled = true;
  newRunBtn.disabled = true;
  setMessage("Running new matching...");

  try {
    const result = await runNewMatchJob();

    confirmedGridApi.setGridOption("rowData", result.confirmed || []);
    newNamesGridApi.setGridOption("rowData",  result.new_names  || []);
    document.querySelector("#add-new-names-btn").disabled  = true;
    document.querySelector("#new-names-message").textContent = "";

    Object.keys(possiblePicks).forEach((k) => delete possiblePicks[k]);
    expandedRows.clear();
    possibleBaseRows = (result.possible || []).slice().sort((a, b) => {
      const bestScore = (row) => {
        const c = row.candidates || [];
        return c.length ? Math.max(...c.map((x) => x.final_score ?? 0)) : -Infinity;
      };
      return bestScore(b) - bestScore(a);
    });
    possibleGridApi.setGridOption("rowData", possibleBaseRows);
    document.querySelector("#write-possible-btn").disabled  = true;
    document.querySelector("#possible-message").textContent = "";

    const meta = result.meta || {};
    setMessage(
      `Done (new). Missing: ${meta.missing_count || 0}, confirmed: ${meta.confirmed_count || 0}, ` +
      `possible: ${meta.possible_count || 0}, new: ${meta.new_names_count || 0}`
    );
  } catch (error) {
    console.error(error);
    setMessage(error.message || "New match run failed.", true);
    alert(error.message || "New match run failed.");
  } finally {
    runBtn.disabled = false;
    newRunBtn.disabled = false;
  }
}

// ── Init ──────────────────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", async () => {
  confirmedGridApi = agGrid.createGrid(
    document.querySelector("#confirmed-grid"), buildConfirmedGrid()
  );
  possibleGridApi = agGrid.createGrid(
    document.querySelector("#possible-grid"),  buildPossibleGrid()
  );
  newNamesGridApi = agGrid.createGrid(
    document.querySelector("#new-names-grid"), buildNewNamesGrid()
  );

  attachTabs();

  document.querySelector("#run-matching").addEventListener("click", runMatching);
  document.querySelector("#run-new-matching").addEventListener("click", runNewMatching);
  document.querySelector("#update-combined").addEventListener("click", async () => {
    const btn = document.querySelector("#update-combined");
    btn.disabled = true;
    setMessage("Updating combined master...");
    try {
      const result = await updateCombinedMaster();
      const rows = result.meta?.rows_written ?? 0;
      setMessage(`Combined master updated: ${result.output_file} (${rows} rows written)`);
    } catch (err) {
      setMessage(err.message || "Update failed.", true);
    } finally {
      btn.disabled = false;
    }
  });
  document.querySelector("#confirm-selected-btn").addEventListener("click", confirmSelected);
  document.querySelector("#write-possible-btn").addEventListener("click", writePossiblePicks);
  document.querySelector("#add-new-names-btn").addEventListener("click", addSelectedNewNames);

  // ── Modal button handlers ──────────────────────────────────────────────────
  document.querySelector("#modal-cancel").addEventListener("click", () => {
    document.querySelector("#new-name-modal").style.display = "none";
  });

  document.querySelector("#modal-back").addEventListener("click", () => {
    saveCurrentModalRow();
    modalIndex--;
    renderModalRow();
  });

  document.querySelector("#modal-next").addEventListener("click", async () => {
    saveCurrentModalRow();
    if (modalIndex < modalRows.length - 1) {
      modalIndex++;
      renderModalRow();
    } else {
      await submitModalRows();
    }
  });

  try {
    const status = await fetchMatchStatus();
    if (status?.last_run?.created_at) {
      setMessage(`Last run: ${status.last_run.created_at}`);
    }
  } catch (_) {
    // status is optional on first load
  }
});

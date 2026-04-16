/* Workflow status panel — loaded on every page */

let _workflowListenersAttached = false;

async function loadWorkflowStatus() {
  try {
    const status = await fetch("/api/workflow/status").then(r => r.json());
    renderWorkflowStatus(status);
  } catch (e) {
    console.warn("Could not load workflow status", e);
  }
}

function renderWorkflowStatus(s) {
  const live = s.live || {};

  // --- Export step ---
  const exp = s.last_export;
  const exportEl = document.querySelector("#ws-export");
  const exportDetail = document.querySelector("#ws-export-detail");
  if (exportEl && exportDetail) {
    if (exp) {
      const dt = new Date(exp.timestamp).toLocaleDateString();
      exportDetail.textContent = `${dt} \u00b7 ${(exp.rows || 0).toLocaleString()} rows`;
      exportEl.className = "workflow-step ok";
    } else {
      exportDetail.textContent = "Not run yet";
      exportEl.className = "workflow-step todo";
    }
  }

  // --- Titles step ---
  const unmapped = live.unmapped_titles || 0;
  const titlesEl = document.querySelector("#ws-titles");
  const titlesDetail = document.querySelector("#ws-titles-detail");
  if (titlesEl && titlesDetail) {
    if (unmapped === 0) {
      titlesDetail.textContent = "All titles mapped";
      titlesEl.className = "workflow-step ok";
    } else {
      titlesDetail.innerHTML = '<a href="/titles" style="color:#f0a500">' + unmapped + " unmapped</a>";
      titlesEl.className = "workflow-step warn";
    }
  }

  // --- Pipeline step ---
  const pipe = s.last_pipeline;
  const pipeEl = document.querySelector("#ws-pipeline");
  const pipeDetail = document.querySelector("#ws-pipeline-detail");
  const pipeBtn = document.querySelector("#ws-run-pipeline");

  const exportTime = exp ? new Date(exp.timestamp) : null;
  const pipeTime = pipe ? new Date(pipe.timestamp) : null;
  const pipeStale = exportTime && (!pipeTime || exportTime > pipeTime);

  if (pipeEl && pipeDetail) {
    if (!pipe) {
      pipeDetail.textContent = "Not run yet";
      pipeEl.className = "workflow-step todo";
      if (pipeBtn) pipeBtn.style.display = "block";
    } else if (pipeStale) {
      const dt = new Date(pipe.timestamp).toLocaleDateString();
      pipeDetail.textContent = dt + " \u00b7 needs update";
      pipeEl.className = "workflow-step warn";
      if (pipeBtn) pipeBtn.style.display = "block";
    } else {
      const dt = new Date(pipe.timestamp).toLocaleDateString();
      pipeDetail.textContent = dt + " \u00b7 " + (pipe.total_rows || 0).toLocaleString() + " rows \u00b7 " + (pipe.gcmid_resolved_pct || 0) + "% resolved";
      pipeEl.className = "workflow-step ok";
      if (pipeBtn) pipeBtn.style.display = "none";
    }
  }

  // --- Match step ---
  const match = s.last_match;
  const matchEl = document.querySelector("#ws-match");
  const matchDetail = document.querySelector("#ws-match-detail");
  const matchBtn = document.querySelector("#ws-run-match");

  const matchStale = pipeTime && (!match || new Date(match.timestamp) < pipeTime);

  if (matchEl && matchDetail) {
    if (!match) {
      matchDetail.textContent = "Not run yet";
      matchEl.className = "workflow-step todo";
      if (matchBtn) matchBtn.style.display = "block";
    } else if (matchStale) {
      matchDetail.textContent = "Pipeline updated \u2014 re-run matching";
      matchEl.className = "workflow-step warn";
      if (matchBtn) matchBtn.style.display = "block";
    } else {
      const dt = new Date(match.timestamp).toLocaleDateString();
      const inbox = (match.possible || 0) + (match.new_names || 0);
      if (inbox > 0) {
        matchDetail.innerHTML = dt + ' \u00b7 <a href="/match" style="color:#f0a500">' + inbox + " need review</a>";
        matchEl.className = "workflow-step warn";
      } else {
        matchDetail.textContent = dt + " \u00b7 all resolved \u2713";
        matchEl.className = "workflow-step ok";
      }
      if (matchBtn) matchBtn.style.display = "none";
    }
  }

  // --- Quick action buttons (attach listeners only once) ---
  if (!_workflowListenersAttached) {
    _workflowListenersAttached = true;

    if (pipeBtn) {
      pipeBtn.addEventListener("click", async () => {
        pipeBtn.disabled = true;
        pipeBtn.textContent = "Running\u2026";
        try {
          await fetch("/api/master/new_combine", { method: "POST" });
          await loadWorkflowStatus();
        } finally {
          pipeBtn.disabled = false;
          pipeBtn.textContent = "\u25B6 Run Pipeline";
        }
      });
    }

    if (matchBtn) {
      matchBtn.addEventListener("click", async () => {
        matchBtn.disabled = true;
        matchBtn.textContent = "Running\u2026";
        try {
          await fetch("/api/match/new_run", { method: "POST" });
          await loadWorkflowStatus();
        } finally {
          matchBtn.disabled = false;
          matchBtn.textContent = "\u25B6 Run Matching";
        }
      });
    }
  }
}

document.addEventListener("DOMContentLoaded", loadWorkflowStatus);

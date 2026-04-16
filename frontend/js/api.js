async function parseError(response, fallback) {
  let detail = "";
  try {
    const payload = await response.json();
    if (payload && payload.detail) {
      detail = typeof payload.detail === "string" ? payload.detail : JSON.stringify(payload.detail);
    }
  } catch (_) {
    detail = await response.text();
  }
  return `${fallback}: ${response.status}${detail ? ` ${detail}` : ""}`;
}

async function fetchUnmappedTitles() {
  const response = await fetch("/api/unmapped_titles");
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to load unmapped titles"));
  }
  return response.json();
}

async function fetchTitleConflicts() {
  const response = await fetch("/api/title_conflicts");
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to load title conflicts"));
  }
  return response.json();
}

async function applyTitleMappings(rows) {
  const response = await fetch("/api/apply_title_mappings", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ rows }),
  });

  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to apply title mappings"));
  }

  return response.json();
}

async function fetchProjects() {
  const response = await fetch("/api/projects");
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to load projects"));
  }
  return response.json();
}

async function saveIncludedProjects(includedProjectIds) {
  const response = await fetch("/api/projects/included", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ included_project_ids: includedProjectIds }),
  });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to save project selection"));
  }
  return response.json();
}

async function fetchManagedProjects() {
  const response = await fetch("/api/projects/managed");
  if (!response.ok) throw new Error(await parseError(response, "Failed to load managed projects"));
  return response.json();
}

async function saveManagedProjects(data) {
  const response = await fetch("/api/projects/managed", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  if (!response.ok) throw new Error(await parseError(response, "Failed to save managed projects"));
  return response.json();
}

async function runLiveExport() {
  const response = await fetch("/api/export/run_live", { method: "POST" });
  if (!response.ok) {
    throw new Error(await parseError(response, "Live export failed"));
  }
  return response.json();
}

async function runFullExport(includedProjectIds) {
  const payload = includedProjectIds ? { included_project_ids: includedProjectIds } : {};
  const response = await fetch("/api/export/run", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to run export"));
  }
  return response.json();
}

async function updateCombinedMaster() {
  const response = await fetch("/api/master/update_combined", {
    method: "POST",
  });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to update combined master"));
  }
  return response.json();
}

async function runNewCombine() {
  const response = await fetch("/api/master/new_combine", { method: "POST" });
  if (!response.ok) {
    throw new Error(await parseError(response, "New combine failed"));
  }
  return response.json();
}

async function runMatchJob() {
  const response = await fetch("/api/match/run", {
    method: "POST",
  });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to run matching"));
  }
  return response.json();
}

async function runNewMatchJob() {
  const response = await fetch("/api/match/new_run", { method: "POST" });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to run new matching"));
  }
  return response.json();
}

async function fetchMatchStatus() {
  const response = await fetch("/api/match/status");
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to load match status"));
  }
  return response.json();
}

async function browseFiles() {
  const response = await fetch("/api/browse/files");
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to list browse files"));
  }
  return response.json();
}

async function browseQuery(payload) {
  const response = await fetch("/api/browse/query", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to query browse data"));
  }
  return response.json();
}

async function browseDistinct(payload) {
  const response = await fetch("/api/browse/distinct", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to load distinct browse values"));
  }
  return response.json();
}

async function addToRegistry(entries) {
  const response = await fetch("/api/match/add_to_registry", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ entries }),
  });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to add new names to Registry"));
  }
  return response.json();
}

async function addNewNames(entries) {
  const response = await fetch("/api/match/add_new_names", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ entries }),
  });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to add new names to Names.xlsx"));
  }
  return response.json();
}

async function runSfIssuesCheck() {
  const response = await fetch("/api/sf_issues/run", { method: "POST" });
  if (!response.ok) throw new Error(await parseError(response, "SF Issues check failed"));
  return response.json();
}

async function saveSfIssuesState(changes) {
  const response = await fetch("/api/sf_issues/save_state", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ changes }),
  });
  if (!response.ok) throw new Error(await parseError(response, "Failed to save state"));
  return response.json();
}

async function exportSfIssues(rows, filename) {
  const response = await fetch("/api/sf_issues/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ rows, filename }),
  });
  if (!response.ok) throw new Error(await parseError(response, "Export failed"));
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename + ".xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

async function browseExport(payload) {
  const response = await fetch("/api/browse/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.detail || `Export failed: ${response.status}`);
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = payload.filename ? `${payload.filename}.xlsx` : "export.xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

async function browseRegistryLookup(gcmids) {
  const response = await fetch("/api/browse/registry_lookup", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ gcmids }),
  });
  if (!response.ok) throw new Error(await parseError(response, "Registry lookup failed"));
  return response.json();
}

async function browseRegistryExport(gcmids, filename) {
  const response = await fetch("/api/browse/registry_export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ gcmids, filename }),
  });
  if (!response.ok) throw new Error(await parseError(response, "Registry export failed"));
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename + ".xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

async function fetchRegistry() {
  const response = await fetch("/api/registry");
  if (!response.ok) throw new Error(await parseError(response, "Failed to load registry"));
  return response.json();
}

async function saveRegistry(changes) {
  const response = await fetch("/api/registry/save", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ changes }),
  });
  if (!response.ok) throw new Error(await parseError(response, "Failed to save registry"));
  return response.json();
}

async function fetchGeneralTitles() {
  const response = await fetch("/api/general_titles");
  if (!response.ok) throw new Error(await parseError(response, "Failed to load general titles"));
  return response.json();
}

async function confirmMatchesToGcmidMap(entries) {
  const response = await fetch("/api/match/confirm_to_gcmid_map", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ entries }),
  });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to write matches to GCMID Map"));
  }
  return response.json();
}

async function confirmMatchesToHelper(entries) {
  const response = await fetch("/api/match/confirm_to_helper", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ entries }),
  });
  if (!response.ok) {
    throw new Error(await parseError(response, "Failed to write matches to Helper"));
  }
  return response.json();
}

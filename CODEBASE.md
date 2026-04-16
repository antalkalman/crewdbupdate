# CrewDB Updater — Codebase Documentation

## What it does

A FastAPI web application for maintaining Pioneer Pictures' internal crew database.
It pulls data from the **CrewCall Crew Manager API**, merges it with historical Excel data,
maps job titles to standardised general titles, and fuzzy-matches unidentified crew members
to the known crew database.

---

## Core Workflow (New Pipeline)

```
1. Export      → Fetch startforms from CrewCall API → SFlist_YYYYMMDD_HHMM.xlsx
2. SF Issues   → Validate SFlist for missing fields, unsigned SFs, no fees
3. Combine     → Merge SF export + historical data → CrewIndex.xlsx (28 cols) + update CrewRegistry.xlsx (16 cols)
4. Title Map   → Map unmapped Title–Project pairs via TitleMap.xlsx (+ Helper.xlsx)
5. Match       → Fuzzy-match unmatched crew → Confirmed / Possible / New Names
6. Confirm     → Write matches to GCMID_Map.xlsx
7. New Names   → Review modal → write to CrewRegistry.xlsx
```

---

## Project Structure

```
crewdbupdate/
│
├── backend/                          # FastAPI application
│   ├── app.py                        # Entry point: routes, static files, page serving
│   ├── config.py                     # Env vars: CM_API_BASE_URL, CM_API_TOKEN
│   ├── dependencies.py               # Singleton service injection (TitleService)
│   │
│   ├── utils/
│   │   ├── paths.py                  # Central path constants (BASE_DIR, MASTER_DIR, NEW_MASTER_DIR, etc.)
│   │   └── config_store.py           # Save/load included_projects.json
│   │
│   ├── services/
│   │   ├── cm_api.py                 # CrewCall API client (paginated GET, Bearer auth)
│   │   ├── export_service.py         # Orchestrates export workflow
│   │   ├── title_service.py          # Title mapping logic (unmapped pairs, TitleMap.xlsx + Helper.xlsx)
│   │   └── match_service.py          # Wrapper around x_new_match.run_matching()
│   │
│   └── routes/
│       ├── health.py                 # GET /api/health
│       ├── workflow.py               # GET /api/workflow/status, update_status() helper
│       ├── export.py                 # GET/POST /api/projects, POST /api/export/run, GET/POST /projects/managed
│       ├── master.py                 # POST /api/master/update_combined, POST /api/master/new_combine
│       ├── titles.py                 # GET /api/unmapped_titles, POST /api/apply_title_mappings, GET /api/general_titles, GET /api/title_conflicts
│       ├── match.py                  # POST /api/match/run|new_run, confirm_to_helper|confirm_to_gcmid_map, add_new_names|add_to_registry
│       ├── browse.py                 # POST /api/browse/query|distinct|preview|export|registry_lookup|registry_export
│       ├── registry.py               # GET /api/registry, POST /api/registry/save
│       └── sf_issues.py              # POST /api/sf_issues/run|save_state|export
│
├── frontend/
│   ├── index.html                    # Redirects to /export
│   ├── export.html                   # Project selection + export trigger
│   ├── sf_browser.html               # Browse SFlist_*.xlsx files
│   ├── sf_issues.html                # SFlist validation checker
│   ├── crew_explorer.html            # CrewIndex browser with tag filters + Export view
│   ├── registry.html                 # Editable CrewRegistry browser
│   ├── combined_browser.html         # (redirects to /crew_explorer)
│   ├── names_browser.html            # Browse Names.xlsx
│   ├── titles.html                   # Title Mapper UI
│   ├── match.html                    # Match Inbox (3 tabs) + New Name review modal
│   ├── styles.css                    # App-wide styles (sidebar, workflow panel, filters, AG Grid)
│   └── js/
│       ├── api.js                    # Central async API client (~25 endpoints)
│       ├── workflow.js               # Sidebar workflow status panel (loaded on all pages)
│       ├── export.js                 # Project grid with checkboxes, withBusy() wrapper
│       ├── sf_browser.js             # File dropdown, pagination, CSV export
│       ├── sf_issues.js              # SF validation grid with state persistence + export
│       ├── crew_explorer.js          # Tag-based filter bar, 4 view modes, GCMID→registry lookup
│       ├── registry.js               # Editable CrewRegistry grid with pending changes
│       ├── title_mapper.js           # TypeAheadGeneralTitleEditor, green row on mapped
│       └── match.js                  # 3 AG Grids in tabs + review modal + "Add as New"
│
├── New_Master_Database/              # Active pipeline data
│   ├── projects.json                 # Three-state project config (live/historical/skip)
│   ├── status.json                   # Workflow state persistence (timestamps + stats)
│   ├── CrewIndex.xlsx                # Combined dataset (26 cols) + 6 derived sheets
│   ├── CrewRegistry.xlsx             # ~6,284 crew members, 14 cols, Status dropdown
│   ├── GCMID_Map.xlsx                # ~17,326 CM-Job → CM ID mappings
│   ├── TitleMap.xlsx                 # Title conv (~11,125) + General Title (511)
│   └── Historical/                   # 35 per-project xlsx files
│
├── Master_database/                  # Old pipeline (hidden but functional)
│   ├── Combined_All_CrewData.xlsx    # Main dataset + derived sheets
│   ├── Names.xlsx                    # Known crew (CM ID, names, contact, title)
│   ├── Helper.xlsx                   # Mapping tables (GCMID, Title conv, General Title, FProjects)
│   └── ...
│
├── SF_Archive/
│   └── SFlist_YYYYMMDD_HHMM.xlsx    # Timestamped exports from CrewCall API
│
├── x_new_combine.py                  # New pipeline: combine + preprocess + CrewRegistry update
├── x_new_match.py                    # New pipeline: fuzzy matching
├── X_Ultimate_full_export_api.py     # Fetches startforms from API → SFlist_*.xlsx
├── X_master_combine_and_preprocess.py  # Old pipeline orchestrator
├── x_master_combined.py              # Old pipeline: merge
├── x_master_preprocess.py            # Old pipeline: derived sheets
├── x_master_match.py                 # Old pipeline: fuzzy matching
├── migrate_*.py                      # One-time migration scripts
│
├── .env.title_mapper                 # API credentials (CM_API_BASE_URL, CM_API_TOKEN)
├── Start crew_db_updater.command     # macOS launcher (starts uvicorn on port 8000)
└── Stop crew_db_updater.command      # macOS killer (stops uvicorn)
```

---

## Backend API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/health` | Liveness check |
| GET | `/api/workflow/status` | Workflow panel data (persisted + live stats) |
| GET | `/api/projects` | List all projects with inclusion status |
| POST | `/api/projects/included` | Save selected project IDs |
| GET | `/api/projects/managed` | Read projects.json directly |
| POST | `/api/projects/managed` | Write projects.json |
| POST | `/api/export/run` | Run full SF export |
| POST | `/api/export/run_live` | Export live projects only (writes status.json) |
| POST | `/api/master/update_combined` | Rebuild Combined_All_CrewData.xlsx (old) |
| POST | `/api/master/new_combine` | Run new pipeline (writes status.json) |
| GET | `/api/unmapped_titles` | Get unmapped Title-Project pairs |
| POST | `/api/apply_title_mappings` | Append new mappings to TitleMap + Helper |
| GET | `/api/general_titles` | List all valid General Titles from TitleMap.xlsx |
| GET | `/api/title_conflicts` | Get title mapping conflicts |
| POST | `/api/match/run` | Run fuzzy matching (old) |
| POST | `/api/match/new_run` | Run new matching (writes status.json) |
| GET | `/api/match/status` | Get last match run status |
| POST | `/api/match/confirm_to_helper` | Write matches to Helper.xlsx GCMID sheet (old) |
| POST | `/api/match/confirm_to_gcmid_map` | Write matches to GCMID_Map.xlsx (new) |
| POST | `/api/match/add_new_names` | Add new names to Names.xlsx (old) |
| POST | `/api/match/add_to_registry` | Add new names to CrewRegistry.xlsx (new, with actual_title_override, status, note) |
| GET | `/api/browse/files` | List browseable Excel files (SFlist, Combined, CrewIndex, Names) |
| POST | `/api/browse/query` | Paginated/filtered query on an Excel file |
| POST | `/api/browse/distinct` | Distinct values for filter dropdowns |
| POST | `/api/browse/preview` | Preview first N rows of a file |
| POST | `/api/browse/export` | Export all filtered rows as xlsx (no pagination limit) |
| POST | `/api/browse/registry_lookup` | Look up CrewRegistry rows by GCMID list |
| POST | `/api/browse/registry_export` | Export CrewRegistry rows by GCMID list as xlsx |
| GET | `/api/registry` | Read all CrewRegistry rows |
| POST | `/api/registry/save` | Save cell-level edits to CrewRegistry |
| POST | `/api/sf_issues/run` | Validate latest SFlist for issues |
| POST | `/api/sf_issues/save_state` | Persist checked/note state |
| POST | `/api/sf_issues/export` | Export selected issues as xlsx (one sheet per project) |

---

## Frontend Pages

| URL | File | Purpose |
|-----|------|---------|
| `/export` | export.html | Select projects, run export, run new pipeline |
| `/sf_browser` | sf_browser.html | Browse any SFlist_*.xlsx with pagination |
| `/sf_issues` | sf_issues.html | Validate SFlist: missing fields, unsigned SFs, no fees |
| `/crew_explorer` | crew_explorer.html | Browse CrewIndex with tag filters, 4 view modes, Excel export |
| `/registry` | registry.html | Edit CrewRegistry: names, titles, status, contacts |
| `/titles` | titles.html | Map unmapped job titles, review dynamic conflicts |
| `/match` | match.html | Review matches in 3 tabs + New Name review modal |
| `/combined_browser` | — | Redirects to `/crew_explorer` |
| `/names_browser` | — | Redirects to `/registry` |

All pages include the **Workflow Status Panel** in the sidebar (js/workflow.js).

---

## CrewRegistry.xlsx — Status Column

CrewRegistry has 16 columns: CM ID, Sure/First/Nick Name, Actual Title, Status, Actual Phone, Actual Email, Note, Last General Title/Department, Title Flag, Last Email/Phone, Shows Worked, Actual Name.

Column F stores crew status as a text dropdown with 4 values:
- **Active** — regular crew member
- **Retired** — no longer working
- **Foreign** — foreign national crew
- **External** — external contractor

**Pipeline behavior:** All statuses are included in combine, derived sheets, and matching. Status is purely a view/filter label for the UI.

**Excel validation:** Inline list `"Active,Retired,Foreign,External"` (not a cross-sheet reference, which openpyxl corrupts).

**ensure_status_validation(ws):** Utility in x_new_combine.py called before every CrewRegistry save to re-apply the inline validation and fix the table column name.

**Row highlighting in CrewIndex.xlsx:** Retired = light red (#FFE0E0), Foreign = light blue (#DDEEFF), External = light grey (#EEEEEE).

---

## Workflow Status Panel

Sidebar panel on every page showing pipeline state:
- **Export** — last run date + row count
- **Titles** — unmapped count (live check, links to /titles)
- **Pipeline** — last run + staleness warning if export is newer
- **Match** — last run + inbox count (links to /match)

Quick action buttons appear when needed (Run Pipeline, Run Matching).

State persisted to `New_Master_Database/status.json`. Backend routes write to it after each action (export, combine, match).

---

## Match Page — Tab Details

### Confirmed tab
- Checkbox column on Source Key (header checkbox = select all)
- **"Write selected to GCMID Map"** button
- Each row: source_key, name, project, job title, suggested GCMID, matched name, DB title, all scores

### Possible tab
- Rows sorted descending by **Best Score**, with Phone + Email columns after score
- **Chevron** column — expand to show candidate detail panel
- **Status** column — "Picked" (green) or "Skipped" (grey)
- Expanded panel: candidate table with Pick / None of these / **Add as New** buttons
- "Add as New" opens the review modal pre-populated with the Possible row's data
- **"Write picks to GCMID Map"** button

### New Names tab
- Grid: source key, name, project, job title, dept, title, phone, email, reason
- Select rows → **"Add selected to Registry"** → opens review modal
- **Review Modal**: step-through editor (1 of N) with:
  - First Name / Sure Name (pre-split from crew list name)
  - Actual Title dropdown (General Titles from TitleMap, pre-filled from match results)
  - Phone / Email (editable, clear if temporary)
  - Status dropdown (Active/Retired/Foreign/External, default Active)
  - Note (free text, only written if user provides it)
  - Back/Next navigation, "Confirm & Write All" on last row
- Writes to CrewRegistry.xlsx via POST /api/match/add_to_registry

---

## SF Issues Page

- Validates latest SFlist for: missing required fields, unsigned startforms past start date, no fees
- BD startforms skip name/contact field checks
- Project filter dropdown + issues-only toggle
- Editable Note column (inline, auto-saves after 1.5s)
- Extra column picker to temporarily add any SFlist column
- Export selected or all issues as xlsx (one sheet per project)
- State persisted to `sf_issues_state.json` (checked status + edited notes survive refresh)

---

## Crew Explorer Page

- Reads CrewIndex.xlsx (28 cols) via browse API
- Unified tag-based filter bar: Name, Title, Department, Project, Status, Origin
- Title picker filters by selected departments via DEPT_TITLE_MAP
- 4 view modes: General, On Project, Full, Export
- Export view: two-step query (CrewIndex filters → unique GCMIDs → CrewRegistry lookup)
- Server-side pagination with numeric-aware sort (Department ID, Title ID)
- Excel export of all filtered rows (no pagination limit)

---

## Registry Page

- Editable AG Grid on CrewRegistry.xlsx (16 cols)
- Editable: Sure/First/Nick Name, Actual Title (dropdown), Status (dropdown), Actual Phone/Email, Note
- Auto columns (greyed out): Last General Title/Department, Title Flag, Last Email/Phone, Shows Worked, Actual Name
- Pending changes tracked by cm_id|field key, saved via POST /api/registry/save
- Default filter: Status = Active

---

## Data Processing Scripts

### x_new_combine.py
- Reads projects.json → Historical files + SFlist → CrewIndex.xlsx (28 cols incl. Department ID + Title ID)
- Updates CrewRegistry.xlsx auto columns with non-destructive writes (never blanks existing values)
- All statuses included in all derived sheets
- Actual Details sheet has 8 cols: Actual + Last phone/email
- ensure_status_validation(ws) called before every save
- BLOCKED_EMAILS filtering applied during source file reading

### x_new_match.py — Fuzzy Matching Algorithm
All statuses included. Actual Phone/Email preferred over Last Phone/Email. Dept matching uses set of all General Departments per GCMID + Actual Title → dept lookup from TitleMap.

Scores each unmatched crew record against all known crew:

| Signal | Method | Max score |
|--------|--------|-----------|
| Name | Tokenised Levenshtein (exact=1.0, prefix=0.75, dist1=0.5, dist2=0.25) | x1.5 weight |
| Email | Exact=1.0, dist1=0.5 | 1.0 |
| Phone | Normalised exact=1.0, dist1=0.5 | 1.0 |
| Department | General Dept match (set of all depts per GCMID)=0.5 | 0.5 |

**Thresholds:**
- **Confirmed**: name_score >= 1.25 AND (email + phone) >= 1.0
- **Possible**: final_score >= 1.25 (top 5 candidates)
- **New Name**: no candidate above threshold

All statuses are matching candidates (Status is view-level only).

---

## Configuration

### Environment variables (`.env.title_mapper`)
```
CM_API_BASE_URL=https://manager.crewcall.hu/api/
CM_API_TOKEN=<bearer token>
```

### Included projects (`backend/config_data/included_projects.json`)
Persists which project IDs are selected for export.

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Web framework | FastAPI + uvicorn |
| Data processing | pandas, openpyxl |
| Fuzzy matching | python-Levenshtein |
| API client | requests |
| Frontend grid | AG Grid Community 32.3.3 |
| Validation | pydantic |
| Server | localhost:8000 |

---

## Running the App

**Start:** Double-click `Start crew_db_updater.command`
- Sources `.env.title_mapper`
- Starts uvicorn on `127.0.0.1:8000`
- Opens browser to `/export`

**Stop:** Double-click `Stop crew_db_updater.command`
- Kills the uvicorn process by PID

---

## Changes Made (vs original)

| Date | Change |
|------|--------|
| 2026-04-01 | Sidebar label changed from "CrewDB" to "Crew DB" (version marker) |
| 2026-04-01 | Export: excludes `draft` and `paused` startforms |
| 2026-04-01 | Title Mapper: added Department column (sourced from SF list `Project department`) |
| 2026-04-01 | Match / Confirmed tab: checkbox row selection + "Write selected to Helper" button |
| 2026-04-01 | Match / Possible tab: expandable full-width detail rows with candidate picker |
| 2026-04-01 | Match / Possible tab: Best Score column, sorted descending |
| 2026-04-01 | New backend endpoint: POST /api/match/confirm_to_helper (writes to Helper.xlsx GCMID sheet) |
| 2026-04-05 | Workflow status panel added to sidebar on all 6 pages (status.json + GET /api/workflow/status) |
| 2026-04-05 | Backend routes write to status.json after export, combine, and match |
| 2026-04-05 | CrewRegistry.xlsx: Retired boolean → Status text column (Active/Retired/Foreign/External) |
| 2026-04-05 | Status is view-level only — all statuses included in pipeline, matching, derived sheets |
| 2026-04-05 | ensure_status_validation(ws) prevents openpyxl corruption of Excel dropdown |
| 2026-04-05 | CrewIndex.xlsx: row highlighting by Status (Retired=red, Foreign=blue, External=grey) |
| 2026-04-05 | New Names tab: review modal replaces direct write (step-through editor with title/status/note) |
| 2026-04-05 | GET /api/general_titles endpoint + fetchGeneralTitles() in api.js |
| 2026-04-05 | add_to_registry: supports actual_title_override, note, status fields from modal |
| 2026-04-05 | Note field: only writes user-provided text (no auto-fill from department) |
| 2026-04-05 | Dept matching: uses set of all General Departments per GCMID (not single value) |
| 2026-04-05 | GET /api/projects/managed simplified: reads projects.json directly, no API enrichment |
| 2026-04-06 | BLOCKED_EMAILS = {"pioneer@crewcall.hu"} — internal email treated as empty throughout |
| 2026-04-06 | Dept matching enriched with Actual Title → dept lookup from TitleMap |
| 2026-04-06 | Names Browser → Registry page (/registry): editable AG Grid on CrewRegistry.xlsx |
| 2026-04-06 | Phone formatting: +CC AA BBBB rest display in Registry and Crew Explorer |
| 2026-04-07 | Combined Browser → Crew Explorer (/crew_explorer): reads CrewIndex.xlsx, floating filters |
| 2026-04-07 | Browse API: POST /browse/export endpoint for filtered Excel downloads |
| 2026-04-07 | Title conflicts tab now dynamic — computed live from CrewIndex, not static file |
| 2026-04-07 | Possible tab: "Add as New" button opens review modal for unmatched crew |
| 2026-04-07 | Possible tab: phone + email columns added after score |
| 2026-04-08 | CrewRegistry.xlsx expanded to 16 cols: Actual Phone (col 7) + Actual Email (col 8) from Names.xlsx |
| 2026-04-08 | x_new_combine.py: non-destructive auto-column updates (never blanks existing values) |
| 2026-04-08 | Actual Details sheet expanded to 8 cols (Actual + Last phone/email) |
| 2026-04-08 | x_new_match.py: Actual Phone/Email preferred over Last Phone/Email in candidate pool |
| 2026-04-08 | CrewIndex.xlsx expanded to 28 cols: Department ID + Title ID sort keys |
| 2026-04-08 | Browse API: numeric-aware sort (coerces string columns to float for correct ordering) |
| 2026-04-09 | Crew Explorer: unified tag-based filter bar (Name, Title, Department, Project, Status, Origin) |
| 2026-04-09 | Browse API: cross_column_in filter operator for multi-column OR matching |
| 2026-04-09 | Export view mode: two-step query (CrewIndex→GCMIDs→CrewRegistry lookup) |
| 2026-04-09 | GET /api/general_titles_with_dept + DEPT_TITLE_MAP for title picker dept filtering |
| 2026-04-09 | Browse API: POST /browse/registry_lookup + registry_export (GCMID-based) |
| 2026-04-10 | SF Issues page (/sf_issues): validates SFlist for missing fields, unsigned SFs, no fees |
| 2026-04-10 | SF Issues: state persistence (sf_issues_state.json), project filter, export by project |
| 2026-04-10 | SF Issues: extra column picker, editable notes, checkbox selection for export |

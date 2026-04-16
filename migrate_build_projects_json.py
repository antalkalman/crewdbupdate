"""One-time migration: build projects.json from FProjects + CrewCall API."""

import json
import os
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
import requests
from dotenv import dotenv_values

HELPER = Path("Master_database/Helper.xlsx")
INCLUDED = Path("backend/config_data/included_projects.json")
OUTPUT_JSON = Path("New_Master_Database/projects.json")
OUTPUT_REVIEW = Path("New_Master_Database/projects_review.txt")
ENV_FILE = Path(".env.title_mapper")


def excel_date(serial):
    if not serial or not isinstance(serial, (int, float)):
        return None
    return (datetime(1899, 12, 30) + timedelta(days=int(serial))).strftime("%Y-%m-%d")


def load_fprojects():
    wb = openpyxl.load_workbook(HELPER, read_only=True, data_only=True)
    ws = wb["FProjects"]
    projects = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = str(row[0]).strip() if row[0] else ""
        start = row[1]
        end = row[2]
        if name:
            projects[name] = {
                "start_date": excel_date(start),
                "end_date": excel_date(end),
            }
    wb.close()
    return projects


def load_api_projects():
    env = dotenv_values(ENV_FILE)
    base_url = env.get("CM_API_BASE_URL", "").strip().rstrip("/")
    token = env.get("CM_API_TOKEN", "").strip()

    if not base_url or not token:
        print("  WARNING: API credentials not found, skipping API fetch")
        return None

    headers = {"Authorization": f"Bearer {token}"}
    all_projects = []
    page = 1

    try:
        while True:
            resp = requests.get(
                f"{base_url}/project",
                headers=headers,
                params={"page": page, "per_page": 200},
                timeout=30,
            )
            resp.raise_for_status()
            body = resp.json()

            data = body.get("data", [])
            if isinstance(data, dict):
                data = data.get("items", [])

            for item in data:
                all_projects.append({
                    "id": int(item["id"]),
                    "name": str(item.get("name", "")).strip(),
                })

            meta = body.get("meta", {})
            if len(data) < 200 or page >= meta.get("last_page", page):
                break
            page += 1

        return all_projects
    except Exception as e:
        print(f"  WARNING: API fetch failed: {e}")
        return None


def load_included_ids():
    if not INCLUDED.exists():
        return set()
    data = json.loads(INCLUDED.read_text())
    return set(data.get("included_project_ids", []))


def main():
    print("Loading FProjects from Helper.xlsx...")
    fprojects = load_fprojects()
    print(f"  {len(fprojects)} projects from FProjects")

    print("Fetching projects from CrewCall API...")
    api_projects = load_api_projects()
    api_available = api_projects is not None

    if api_available:
        print(f"  {len(api_projects)} projects from API")
    else:
        print("  API unavailable — all projects will be marked 'historical'")

    included_ids = load_included_ids()
    print(f"  {len(included_ids)} included project IDs")

    # Build lookup
    api_by_name = {}
    api_ids = set()
    if api_available:
        for p in api_projects:
            api_by_name[p["name"]] = p["id"]
            api_ids.add(p["id"])

    # Build project list
    result = []
    seen_names = set()

    # Process API projects first
    if api_available:
        for p in api_projects:
            pid = p["id"]
            name = p["name"]
            dates = fprojects.get(name, {})
            state = "live" if pid in included_ids else "skip"

            result.append({
                "id": pid,
                "name": name,
                "start_date": dates.get("start_date"),
                "end_date": dates.get("end_date"),
                "state": state,
            })
            seen_names.add(name)

    # Add FProjects not in API
    for name, dates in fprojects.items():
        if name not in seen_names:
            api_id = api_by_name.get(name)
            result.append({
                "id": api_id,
                "name": name,
                "start_date": dates.get("start_date"),
                "end_date": dates.get("end_date"),
                "state": "historical",
            })

    # Sort by start_date descending (None last)
    result.sort(key=lambda p: p.get("start_date") or "0000-00-00", reverse=True)

    # Write JSON
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON.write_text(json.dumps({"projects": result}, indent=2, ensure_ascii=False))

    # Count by state
    by_state = {"live": [], "historical": [], "skip": []}
    for p in result:
        by_state[p["state"]].append(p)

    # Write review file
    lines = ["projects.json\n"]
    for state, label in [("live", "LIVE (fetch from API)"), ("historical", "HISTORICAL (read from Historical/ folder)"), ("skip", "SKIP (known but excluded)")]:
        lines.append(f"\n{label}:")
        for p in by_state[state]:
            pid = f"[{p['id']}]" if p["id"] else "[?]"
            sd = p["start_date"] or "?"
            ed = p["end_date"] or "?"
            lines.append(f"  {pid:<8}{p['name']:<30}{sd} -> {ed}")
        if not by_state[state]:
            lines.append("  (none)")

    total_line = f"\nTotal: {len(by_state['live'])} live, {len(by_state['historical'])} historical, {len(by_state['skip'])} skip"
    lines.append(total_line)
    lines.append(f"\nReview and edit {OUTPUT_JSON} manually if any states are wrong.")
    lines.append("Then rename 'historical' projects to match their Historical/ filenames exactly.")

    review_text = "\n".join(lines)
    OUTPUT_REVIEW.write_text(review_text)

    # Print summary
    print(f"\n{review_text}")
    print(f"\nOutput: {OUTPUT_JSON}")


if __name__ == "__main__":
    main()

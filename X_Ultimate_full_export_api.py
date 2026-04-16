from __future__ import annotations

import ast
import os
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests


def _fetch_data(base_url: str, token: str, endpoint: str, retries: int = 3, delay: int = 5) -> list[dict]:
    url = f"{base_url.rstrip('/')}/{endpoint.lstrip('/')}"
    headers = {"Authorization": f"Bearer {token}"}

    for attempt in range(retries):
        try:
            response = requests.get(url=url, headers=headers, timeout=60)
            response.raise_for_status()
            payload = response.json()
            data = payload.get("data", []) if isinstance(payload, dict) else []
            if isinstance(data, list):
                return data
            return []
        except requests.exceptions.RequestException as exc:
            print(f"[export] {endpoint} attempt {attempt + 1} failed: {exc}")
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                raise
    return []


def _extract_others(entry_list, label: str) -> dict:
    result: dict[str, str] = {}
    try:
        if isinstance(entry_list, list):
            for i in range(min(2, len(entry_list))):
                item = entry_list[i] or {}
                result[f"{label} {i + 1} description"] = str(item.get("name", "") or "")
                result[f"{label} {i + 1} price"] = str(item.get("price", "") or "")
                result[f"{label} {i + 1} account code"] = str(item.get("account_code", "") or "")
        else:
            parsed = ast.literal_eval(entry_list) if isinstance(entry_list, str) else []
            return _extract_others(parsed, label)
    except Exception as exc:
        result[f"{label} 1 description"] = f"error: {exc}"
    return result


def run_full_export(
    included_project_ids: list[int] | None = None,
    cm_api_base_url: str | None = None,
    cm_api_token: str | None = None,
    base_dir: Path | str | None = None,
) -> dict:
    base_path = Path(base_dir) if base_dir else Path(__file__).resolve().parent
    sf_archive_dir = base_path / "SF_Archive"
    mapping_file = sf_archive_dir / "export_field_mapping.xlsx"

    sf_archive_dir.mkdir(parents=True, exist_ok=True)
    if not mapping_file.exists():
        raise FileNotFoundError(f"Missing mapping file: {mapping_file}")

    api_base_url = (cm_api_base_url or os.getenv("CM_API_BASE_URL", "")).strip()
    api_token = (cm_api_token or os.getenv("CM_API_TOKEN", "")).strip()
    if not api_base_url:
        raise ValueError("Missing CM_API_BASE_URL. Set environment variable CM_API_BASE_URL.")
    if not api_token:
        raise ValueError("Missing CM_API_TOKEN. Set environment variable CM_API_TOKEN.")

    included_ids: set[int] | None
    if included_project_ids is None:
        included_ids = None
    else:
        included_ids = {int(value) for value in included_project_ids}

    print(f"[export] Base directory: {base_path}")
    print(f"[export] SF archive: {sf_archive_dir}")
    print(f"[export] Included project ids: {sorted(included_ids) if included_ids is not None else 'ALL'}")

    projects = _fetch_data(api_base_url, api_token, "project")
    departments = _fetch_data(api_base_url, api_token, "department")
    job_titles = _fetch_data(api_base_url, api_token, "job_title")
    startforms = _fetch_data(api_base_url, api_token, "startform")
    users = _fetch_data(api_base_url, api_token, "user")
    overtimes = _fetch_data(api_base_url, api_token, "overtime")
    templates = _fetch_data(api_base_url, api_token, "startform_template")
    turnarounds = _fetch_data(api_base_url, api_token, "turnaround")
    units = _fetch_data(api_base_url, api_token, "unit")
    working_hours = _fetch_data(api_base_url, api_token, "working_hour")

    project_lookup = {p.get("id"): p.get("name", "") for p in projects}
    department_lookup = {d.get("id"): d.get("name", "") for d in departments}
    job_title_lookup = {j.get("id"): j.get("name", "") for j in job_titles}
    department_sort_lookup = {d.get("id"): d.get("sort", 0) for d in departments}
    job_title_sort_lookup = {j.get("id"): j.get("sort", 0) for j in job_titles}
    overtime_lookup = {o.get("id"): o.get("name", "") for o in overtimes}
    turnaround_lookup = {t.get("id"): t.get("name", "") for t in turnarounds}
    unit_lookup = {u.get("id"): u.get("name", "") for u in units}
    working_hour_lookup = {w.get("id"): w.get("name", "") for w in working_hours}
    template_lookup = {t.get("id"): t.get("title", "") for t in templates}
    unit_sort_lookup = {u.get("id"): u.get("sort", 0) for u in units}

    df_mapping = pd.read_excel(mapping_file)
    df_sf = pd.json_normalize(startforms)

    if included_ids is not None and "project_id" in df_sf.columns:
        df_sf = df_sf[df_sf["project_id"].isin(included_ids)]

    # Exclude draft and paused startforms
    if "state" in df_sf.columns:
        excluded_states = {"draft", "paused"}
        before = len(df_sf)
        df_sf = df_sf[~df_sf["state"].str.lower().str.strip().isin(excluded_states)]
        print(f"[export] Excluded {before - len(df_sf)} startforms with state in {excluded_states}")

    user_lookup = {
        u.get("id"): {
            "User name": u.get("name", ""),
            "User surname": u.get("surname", ""),
            "User email": u.get("email", ""),
            "User phone": u.get("phone", ""),
        }
        for u in users
    }

    if "user_id" in df_sf.columns:
        user_df = df_sf["user_id"].map(user_lookup).apply(pd.Series)
        df_sf = pd.concat([df_sf, user_df], axis=1)

    for col in df_sf.columns:
        if isinstance(col, str) and "date" in col.lower():
            try:
                df_sf[col] = pd.to_datetime(df_sf[col], errors="coerce").dt.tz_localize(None)
            except Exception:
                pass

    if "project_id" in df_sf.columns:
        df_sf["Project"] = df_sf["project_id"].map(project_lookup)
    if "project_department_id" in df_sf.columns:
        df_sf["Project department"] = df_sf["project_department_id"].map(department_lookup)
        df_sf["Dept Sort"] = df_sf["project_department_id"].map(department_sort_lookup)
    if "project_job_title_id" in df_sf.columns:
        df_sf["Project job title"] = df_sf["project_job_title_id"].map(job_title_lookup)
        df_sf["Title Sort"] = df_sf["project_job_title_id"].map(job_title_sort_lookup)
    if "project_overtime_id" in df_sf.columns:
        df_sf["Project overtime"] = df_sf["project_overtime_id"].map(overtime_lookup)
    if "project_turnaround_id" in df_sf.columns:
        df_sf["Project turnaround"] = df_sf["project_turnaround_id"].map(turnaround_lookup)
    if "project_unit_id" in df_sf.columns:
        df_sf["Project unit"] = df_sf["project_unit_id"].map(unit_lookup)
        df_sf["Unit Sort"] = df_sf["project_unit_id"].map(unit_sort_lookup)
    if "project_working_hour_id" in df_sf.columns:
        df_sf["Project working hour"] = df_sf["project_working_hour_id"].map(working_hour_lookup)
    if "project_startform_id" in df_sf.columns:
        df_sf["Startform template"] = df_sf["project_startform_id"].map(template_lookup)

    if "deal_notes" in df_sf.columns:
        df_sf["deal_notes"] = (
            df_sf["deal_notes"]
            .fillna("")
            .astype(str)
            .str.replace("\n", " ", regex=False)
            .str.replace("\r", " ", regex=False)
            .str.replace('"', "", regex=False)
            .str.replace(",", ";", regex=False)
            .str.strip()
        )

    if "type" in df_sf.columns and "sf_number" in df_sf.columns:
        df_sf["Sf number"] = df_sf["type"].fillna("").astype(str) + df_sf["sf_number"].astype(str)

    if "crew_member_id" in df_sf.columns:
        df_sf["Crew member id"] = df_sf["crew_member_id"].apply(
            lambda x: f"CM{int(x)}" if pd.notnull(x) else ""
        )

    if "daily_others" in df_sf.columns:
        df_daily = df_sf["daily_others"].apply(lambda x: pd.Series(_extract_others(x, "Daily others")))
    else:
        df_daily = pd.DataFrame(index=df_sf.index)

    if "weekly_others" in df_sf.columns:
        df_weekly = df_sf["weekly_others"].apply(lambda x: pd.Series(_extract_others(x, "Weekly others")))
    else:
        df_weekly = pd.DataFrame(index=df_sf.index)

    if "fee_others" in df_sf.columns:
        df_fee = df_sf["fee_others"].apply(lambda x: pd.Series(_extract_others(x, "Fee others")))
    else:
        df_fee = pd.DataFrame(index=df_sf.index)

    df_sf = pd.concat([df_sf, df_daily, df_weekly, df_fee], axis=1)

    output_dict: dict[str, pd.Series | str] = {}

    for _, row in df_mapping.iterrows():
        export_col = row["Export Column (Official)"]
        api_field = row["API Field (from CrewManager_StartForms.csv)"]

        if export_col in df_sf.columns and (pd.isna(api_field) or str(api_field).strip() == ""):
            output_dict[export_col] = df_sf[export_col]
        elif pd.isna(api_field) or str(api_field).strip() == "":
            output_dict[export_col] = ""
        elif " + " in str(api_field):
            try:
                chunks = [piece.strip() for piece in str(api_field).split("+")]
                if len(chunks) >= 2 and chunks[0] in df_sf.columns and chunks[1] in df_sf.columns:
                    output_dict[export_col] = df_sf[chunks[0]].astype(str) + df_sf[chunks[1]].astype(str)
                else:
                    output_dict[export_col] = ""
            except Exception:
                output_dict[export_col] = ""
        elif str(api_field).startswith('"CM" +'):
            field = str(api_field).split("+", 1)[1].strip()
            if field in df_sf.columns:
                output_dict[export_col] = df_sf[field].apply(
                    lambda x: f"CM{int(x)}" if pd.notnull(x) else ""
                )
            else:
                output_dict[export_col] = ""
        elif str(api_field) in df_sf.columns:
            output_dict[export_col] = df_sf[str(api_field)]
        else:
            output_dict[export_col] = ""

    output_data = pd.DataFrame(output_dict)

    cols_to_move = [
        "is_internal",
        "invite_date",
        "sort_order",
        "created_at",
        "updated_at",
        "deleted_at",
        "downloaded_at",
    ]
    kept_cols = [col for col in output_data.columns if col not in cols_to_move]
    tail_cols = [col for col in cols_to_move if col in output_data.columns]
    output_data = output_data[kept_cols + tail_cols]

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_file = sf_archive_dir / f"SFlist_{timestamp}.xlsx"
    output_data.to_excel(output_file, index=False)

    print(f"[export] Final formatted Excel saved: {output_file}")
    return {
        "ok": True,
        "output_file": output_file.name,
        "output_path": str(output_file),
        "rows": int(len(output_data)),
    }


def _load_saved_project_ids() -> list[int] | None:
    try:
        from backend.utils.config_store import load_included_projects

        ids = sorted(load_included_projects())
        return ids
    except Exception:
        return None


if __name__ == "__main__":
    saved_ids = _load_saved_project_ids()
    result = run_full_export(included_project_ids=saved_ids)
    print(result)

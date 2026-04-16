from __future__ import annotations

import json

from fastapi import HTTPException

from backend.services.cm_api import CrewManagerAPI
from backend.utils.config_store import load_included_projects, save_included_projects
from backend.utils.paths import PROJECTS_JSON
from X_Ultimate_full_export_api import run_full_export


class ExportService:
    def __init__(self) -> None:
        self.cm_api = CrewManagerAPI()

    @staticmethod
    def _parse_date(val) -> str | None:
        if not val:
            return None
        return str(val)[:10]

    def list_projects(self) -> dict:
        projects_raw = self.cm_api.list_projects()
        projects = []
        for row in projects_raw:
            try:
                project_id = int(row.get("id"))
            except (TypeError, ValueError):
                continue
            projects.append({
                "id": project_id,
                "name": str(row.get("name", "")).strip(),
                "start_date": self._parse_date(row.get("start_date")),
                "end_date": self._parse_date(row.get("end_date")),
            })

        included = sorted(load_included_projects())
        return {"projects": projects, "included_project_ids": included}

    def save_included_projects(self, ids: list[int]) -> dict[str, int | bool]:
        normalized = []
        for value in ids:
            try:
                normalized.append(int(value))
            except (TypeError, ValueError):
                continue
        save_included_projects(normalized)
        return {"saved": True, "count": len(set(normalized))}

    def run_export(self, ids: list[int] | None = None) -> dict:
        if ids is None:
            included_ids = sorted(load_included_projects())
        else:
            included_ids = []
            for value in ids:
                try:
                    included_ids.append(int(value))
                except (TypeError, ValueError):
                    continue
            included_ids = sorted(set(included_ids))

        try:
            result = run_full_export(included_project_ids=included_ids)
        except HTTPException:
            raise
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Export failed: {exc}") from exc

        return {
            "ok": True,
            "output_file": result["output_file"],
            "rows": result["rows"],
        }

    def run_live_export(self) -> dict:
        if not PROJECTS_JSON.exists():
            raise HTTPException(status_code=404, detail="projects.json not found")

        data = json.loads(PROJECTS_JSON.read_text())
        live_ids = [
            int(p["id"])
            for p in data.get("projects", [])
            if p.get("state") == "live" and p.get("id") is not None
        ]

        if not live_ids:
            raise HTTPException(status_code=400, detail="No live projects with CM IDs found in projects.json")

        try:
            result = run_full_export(included_project_ids=live_ids)
        except HTTPException:
            raise
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Live export failed: {exc}") from exc

        return {
            "ok": True,
            "output_file": result["output_file"],
            "rows": result["rows"],
            "live_project_ids": live_ids,
            "live_project_count": len(live_ids),
        }

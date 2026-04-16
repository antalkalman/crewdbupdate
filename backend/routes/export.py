from __future__ import annotations

import json

from pydantic import BaseModel
from fastapi import APIRouter, Depends, HTTPException

from backend.services.export_service import ExportService
from backend.utils.paths import NEW_MASTER_DIR, PROJECTS_JSON

router = APIRouter()


class IncludedProjectsPayload(BaseModel):
    included_project_ids: list[int]


class RunExportPayload(BaseModel):
    included_project_ids: list[int] | None = None


def get_export_service() -> ExportService:
    return ExportService()


@router.get("/projects")
def get_projects(service: ExportService = Depends(get_export_service)) -> dict:
    return service.list_projects()


@router.post("/projects/included")
def save_projects_included(
    payload: IncludedProjectsPayload,
    service: ExportService = Depends(get_export_service),
) -> dict[str, int | bool]:
    return service.save_included_projects(payload.included_project_ids)


@router.post("/export/run")
def run_export(
    payload: RunExportPayload | None = None,
    service: ExportService = Depends(get_export_service),
) -> dict:
    return service.run_export(payload.included_project_ids if payload else None)


@router.post("/export/run_live")
def run_live_export(service: ExportService = Depends(get_export_service)) -> dict:
    result = service.run_live_export()

    try:
        from datetime import datetime
        from backend.routes.workflow import update_status
        update_status("last_export", {
            "timestamp": datetime.now().isoformat(),
            "filename": result.get("output_file", ""),
            "live_project_count": result.get("live_project_count", 0),
            "rows": result.get("rows", 0),
        })
    except Exception:
        pass

    return result


@router.get("/projects/managed")
def get_managed_projects() -> dict:
    if not PROJECTS_JSON.exists():
        raise HTTPException(status_code=404, detail="projects.json not found")
    return json.loads(PROJECTS_JSON.read_text())


@router.post("/projects/managed")
def save_managed_projects(payload: dict) -> dict:
    projects_path = NEW_MASTER_DIR / "projects.json"
    projects_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False))
    return {"saved": True, "count": len(payload.get("projects", []))}

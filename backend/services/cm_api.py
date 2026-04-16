from __future__ import annotations

from typing import Any

import requests
from fastapi import HTTPException

from backend.config import CM_API_BASE_URL, CM_API_TOKEN


class CrewManagerAPI:
    def __init__(self, base_url: str | None = None, token: str | None = None) -> None:
        self.base_url = (base_url or CM_API_BASE_URL).strip()
        self.token = (token or CM_API_TOKEN).strip()

    def _validate_config(self) -> None:
        missing = []
        if not self.base_url:
            missing.append("CM_API_BASE_URL")
        if not self.token:
            missing.append("CM_API_TOKEN")
        if missing:
            raise HTTPException(
                status_code=500,
                detail=(
                    "Missing Crew Manager API configuration. Set environment variable(s): "
                    + ", ".join(missing)
                ),
            )

    def _request(self, endpoint: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        self._validate_config()
        url = f"{self.base_url.rstrip('/')}/{endpoint.lstrip('/')}"
        headers = {"Authorization": f"Bearer {self.token}"}
        try:
            response = requests.get(url, headers=headers, params=params or {}, timeout=30)
            response.raise_for_status()
        except requests.exceptions.RequestException as exc:
            raise HTTPException(
                status_code=502,
                detail=f"Crew Manager API request failed for '{endpoint}': {exc}",
            ) from exc

        try:
            return response.json()
        except ValueError as exc:
            raise HTTPException(
                status_code=502,
                detail=f"Crew Manager API returned invalid JSON for '{endpoint}'.",
            ) from exc

    def list_projects(self) -> list[dict]:
        projects: list[dict] = []
        page = 1
        seen_ids: set[int] = set()

        while True:
            payload = self._request("project", params={"page": page, "per_page": 200})
            data = payload.get("data", [])

            if isinstance(data, dict):
                if "items" in data and isinstance(data["items"], list):
                    items = data["items"]
                else:
                    items = []
            elif isinstance(data, list):
                items = data
            else:
                items = []

            if not items:
                if page == 1:
                    print("[CrewManagerAPI] No projects returned from API.")
                break

            new_count = 0
            for item in items:
                if not isinstance(item, dict):
                    continue
                try:
                    project_id = int(item.get("id"))
                except (TypeError, ValueError):
                    continue
                if project_id in seen_ids:
                    continue
                seen_ids.add(project_id)
                projects.append(item)
                new_count += 1

            print(f"[CrewManagerAPI] fetched page {page} project rows: {len(items)}, new: {new_count}")

            meta = payload.get("meta") if isinstance(payload, dict) else None
            last_page = None
            current_page = None
            if isinstance(meta, dict):
                try:
                    last_page = int(meta.get("last_page")) if meta.get("last_page") is not None else None
                except (TypeError, ValueError):
                    last_page = None
                try:
                    current_page = int(meta.get("current_page")) if meta.get("current_page") is not None else None
                except (TypeError, ValueError):
                    current_page = None

            if last_page is not None and current_page is not None:
                if current_page >= last_page:
                    break
            elif len(items) < 200:
                break

            page += 1
            if page > 1000:
                raise HTTPException(status_code=502, detail="Project pagination exceeded safe page limit.")

        projects = sorted(projects, key=lambda p: str(p.get("name", "")).lower())
        return projects

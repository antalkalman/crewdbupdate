from __future__ import annotations

import json
from datetime import datetime, timezone

from backend.utils.paths import CONFIG_DATA_DIR

CONFIG_PATH = CONFIG_DATA_DIR / "included_projects.json"


def _ensure_file() -> None:
    CONFIG_DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(
            json.dumps({"included_project_ids": [], "updated_at": ""}, indent=2),
            encoding="utf-8",
        )


def load_included_projects() -> set[int]:
    _ensure_file()
    raw = CONFIG_PATH.read_text(encoding="utf-8")
    payload = json.loads(raw or "{}")
    values = payload.get("included_project_ids", [])
    result: set[int] = set()
    for value in values:
        try:
            result.add(int(value))
        except (TypeError, ValueError):
            continue
    return result


def save_included_projects(ids: list[int]) -> None:
    _ensure_file()
    unique_sorted = sorted({int(value) for value in ids})
    payload = {
        "included_project_ids": unique_sorted,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    CONFIG_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")

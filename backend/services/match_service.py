from __future__ import annotations

from fastapi import HTTPException

from x_master_match import run_matching

LAST_MATCH_META: dict | None = None


class MatchService:
    def run(self) -> dict:
        global LAST_MATCH_META
        try:
            result = run_matching(write_output=False)
        except FileNotFoundError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Matching failed: {exc}") from exc

        LAST_MATCH_META = result.get("meta", {})
        return result

    def run_new(self) -> dict:
        global LAST_MATCH_META
        try:
            import x_new_match
            result = x_new_match.run_matching(write_output=False)
        except FileNotFoundError as exc:
            raise HTTPException(status_code=404, detail=str(exc)) from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"New matching failed: {exc}") from exc

        LAST_MATCH_META = result.get("meta", {})
        return result

    def status(self) -> dict:
        return {
            "last_run": LAST_MATCH_META,
            "has_run": LAST_MATCH_META is not None,
        }

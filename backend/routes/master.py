from __future__ import annotations

from fastapi import APIRouter, HTTPException

from X_master_combine_and_preprocess import run_master_combine_and_preprocess

router = APIRouter()


@router.post("/master/update_combined")
def update_combined() -> dict:
    try:
        result = run_master_combine_and_preprocess()
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except PermissionError as exc:
        raise HTTPException(
            status_code=423,
            detail=f"{exc} Close the Excel file and retry.",
        ) from exc
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Master update failed: {exc}") from exc

    return {
        "ok": True,
        "output_file": "Master_database/Combined_All_CrewData.xlsx",
        "created_at": result.get("created_at"),
        "meta": result.get("meta", {}),
    }


@router.post("/master/new_combine")
def new_combine() -> dict:
    try:
        from x_new_combine import main as run_new_combine
        result = run_new_combine()
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except PermissionError as exc:
        raise HTTPException(
            status_code=423,
            detail=f"{exc} Close the Excel file and retry.",
        ) from exc
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"New combine failed: {exc}") from exc

    try:
        from backend.routes.workflow import update_status
        meta = result.get("meta", {})
        total = max(meta.get("total_rows", 1), 1)
        update_status("last_pipeline", {
            "timestamp": result.get("created_at"),
            "total_rows": meta.get("total_rows"),
            "gcmid_resolved": meta.get("gcmid_resolved"),
            "gcmid_resolved_pct": round(
                meta.get("gcmid_resolved", 0) / total * 100, 1
            ),
            "title_mapped_pct": round(
                meta.get("title_mapped", 0) / total * 100, 1
            ),
            "warnings": meta.get("warnings", []),
        })
    except Exception:
        pass

    return result

from __future__ import annotations

import io
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal

import pandas as pd
from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field

from backend.utils.paths import BASE_DIR, CREW_REGISTRY_FILE, MASTER_DIR, NEW_MASTER_DIR, SF_ARCHIVE_DIR

router = APIRouter()

# TODO: Replace per-request full sheet loads with chunked/streaming read strategy for very large files.
_DF_CACHE: dict[tuple[str, str, float], tuple[pd.DataFrame, str]] = {}
_CACHE_LIMIT = 8


class SortSpec(BaseModel):
    col: str
    dir: Literal["asc", "desc"] = "asc"


class FilterSpec(BaseModel):
    col: str
    op: Literal["contains", "equals", "in", "gte_date", "contains_normalized", "contains_any", "cross_column_in"]
    value: str | list[str] = ""


class BrowseQueryRequest(BaseModel):
    file_path: str
    sheet: str = ""
    page: int = Field(default=1, ge=1)
    page_size: int = Field(default=200, ge=1, le=2000)
    sort: list[SortSpec] = Field(default_factory=list)
    filters: list[FilterSpec] = Field(default_factory=list)


class BrowsePreviewRequest(BaseModel):
    file_path: str
    sheet: str = ""
    limit: int = Field(default=20, ge=1, le=200)


class BrowseDistinctRequest(BaseModel):
    file_path: str
    sheet: str = ""
    columns: list[str] = Field(default_factory=list)
    filters: list[FilterSpec] = Field(default_factory=list)
    limit: int = Field(default=1000, ge=1, le=5000)


def _to_iso(ts: float) -> str:
    return datetime.fromtimestamp(ts, tz=timezone.utc).isoformat()


def _relative(path: Path) -> str:
    return path.resolve().relative_to(BASE_DIR.resolve()).as_posix()


def _validate_file_path(file_path: str) -> Path:
    candidate = (BASE_DIR / file_path).resolve()
    base = BASE_DIR.resolve()

    try:
        candidate.relative_to(base)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="file_path is outside allowed project folder") from exc

    if not candidate.exists() or not candidate.is_file():
        raise HTTPException(status_code=404, detail=f"File not found: {file_path}")

    if candidate.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise HTTPException(status_code=400, detail="Only Excel files are supported")

    return candidate


def _read_sheet_cached(file_path: Path, sheet: str) -> tuple[pd.DataFrame, str]:
    stat = file_path.stat()
    desired_sheet = (sheet or "").strip()
    key = (str(file_path), desired_sheet, stat.st_mtime)

    if key in _DF_CACHE:
        cached_df, cached_sheet = _DF_CACHE[key]
        return cached_df, cached_sheet

    try:
        workbook = pd.ExcelFile(file_path, engine="openpyxl")
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Could not open workbook: {exc}") from exc

    if not desired_sheet:
        desired_sheet = workbook.sheet_names[0]

    if desired_sheet not in workbook.sheet_names:
        raise HTTPException(status_code=400, detail=f"Sheet '{desired_sheet}' not found in workbook")

    try:
        df = pd.read_excel(file_path, sheet_name=desired_sheet, dtype=str, engine="openpyxl")
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Could not read sheet '{desired_sheet}': {exc}") from exc

    df = df.fillna("")

    if len(_DF_CACHE) >= _CACHE_LIMIT:
        # drop an arbitrary old entry to keep memory bounded in v1
        _DF_CACHE.pop(next(iter(_DF_CACHE)))
    _DF_CACHE[key] = (df, desired_sheet)
    return df, desired_sheet


def _apply_filters(df: pd.DataFrame, filters: list[FilterSpec]) -> pd.DataFrame:
    out = df
    for f in filters:
        if f.op not in ("contains_any", "cross_column_in") and f.col not in out.columns:
            continue
        value_obj: Any = f.value
        series = out[f.col].fillna("").astype(str) if f.col in out.columns else pd.Series([], dtype=str)
        if f.op == "contains":
            value = str(value_obj or "")
            out = out[series.str.contains(value, case=False, na=False)]
        elif f.op == "equals":
            value = str(value_obj or "")
            out = out[series == value]
        elif f.op == "in":
            values = value_obj if isinstance(value_obj, list) else [str(value_obj)]
            values_set = {str(v) for v in values}
            out = out[series.isin(values_set)]
        elif f.op == "gte_date":
            value = str(value_obj or "")
            rhs = pd.to_datetime(value, errors="coerce")
            # Handle mixed date representations: ISO strings, localized strings, and Excel serial numbers.
            lhs = pd.to_datetime(series, errors="coerce")
            numeric = pd.to_numeric(series, errors="coerce")
            lhs_numeric = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
            lhs = lhs.fillna(lhs_numeric)
            if pd.notna(rhs):
                out = out[lhs >= rhs]
        elif f.op == "contains_normalized":
            value = str(value_obj or "")
            lhs = (
                series.str.normalize("NFD")
                .str.replace(r"[\u0300-\u036f]", "", regex=True)
                .str.lower()
                .str.replace(r"[^a-z0-9]", "", regex=True)
            )
            rhs = (
                pd.Series([value])
                .str.normalize("NFD")
                .str.replace(r"[\u0300-\u036f]", "", regex=True)
                .str.lower()
                .str.replace(r"[^a-z0-9]", "", regex=True)
                .iloc[0]
            )
            out = out[lhs.str.contains(rhs, na=False)]
        elif f.op == "contains_any":
            value = str(value_obj or "")
            rhs = (
                pd.Series([value])
                .str.normalize("NFD")
                .str.replace(r"[\u0300-\u036f]", "", regex=True)
                .str.lower()
                .str.replace(r"[^a-z0-9]", "", regex=True)
                .iloc[0]
            )
            if not rhs:
                continue

            mask = pd.Series(False, index=out.index)
            for any_col in out.columns:
                normalized_col = (
                    out[any_col]
                    .fillna("")
                    .astype(str)
                    .str.normalize("NFD")
                    .str.replace(r"[\u0300-\u036f]", "", regex=True)
                    .str.lower()
                    .str.replace(r"[^a-z0-9]", "", regex=True)
                )
                mask = mask | normalized_col.str.contains(rhs, na=False)
            out = out[mask]
        elif f.op == "cross_column_in":
            cols_to_check = [c.strip() for c in f.col.split("|") if c.strip() in out.columns]
            values = value_obj if isinstance(value_obj, list) else [str(value_obj)]
            values_set = {str(v) for v in values}
            if cols_to_check:
                mask = pd.Series(False, index=out.index)
                for check_col in cols_to_check:
                    series = out[check_col].fillna("").astype(str)
                    mask = mask | series.isin(values_set)
                out = out[mask]
    return out


def _apply_sort(df: pd.DataFrame, sort: list[SortSpec]) -> pd.DataFrame:
    if not sort:
        return df

    cols = [s.col for s in sort if s.col in df.columns]
    if not cols:
        return df

    ascending = [s.dir == "asc" for s in sort if s.col in df.columns]

    # Create sort keys — coerce numeric-looking columns to float for correct ordering
    sort_df = df.copy()
    for col in cols:
        numeric = pd.to_numeric(sort_df[col], errors="coerce")
        if numeric.notna().any():
            sort_df[f"__sort_{col}"] = numeric

    sort_keys = [f"__sort_{c}" if f"__sort_{c}" in sort_df.columns else c for c in cols]
    sorted_df = sort_df.sort_values(by=sort_keys, ascending=ascending, kind="mergesort", na_position="last")

    # Drop temporary sort columns and return with original index order
    return sorted_df.drop(columns=[c for c in sorted_df.columns if c.startswith("__sort_")])


def _paginate(df: pd.DataFrame, page: int, page_size: int) -> tuple[pd.DataFrame, int]:
    total_rows = int(len(df))
    start = (page - 1) * page_size
    end = start + page_size
    if start >= total_rows:
        return df.iloc[0:0], total_rows
    return df.iloc[start:end], total_rows


@router.get("/browse/files")
def browse_files() -> dict:
    sf_files = []
    if SF_ARCHIVE_DIR.exists():
        for file in sorted(SF_ARCHIVE_DIR.glob("SFlist_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True):
            sf_files.append(
                {
                    "name": file.name,
                    "path": _relative(file),
                    "mtime": _to_iso(file.stat().st_mtime),
                }
            )

    combined_files = []
    combined = MASTER_DIR / "Combined_All_CrewData.xlsx"
    if combined.exists():
        combined_files.append(
            {
                "name": combined.name,
                "path": _relative(combined),
                "mtime": _to_iso(combined.stat().st_mtime),
            }
        )

    crewindex_files = []
    crewindex = NEW_MASTER_DIR / "CrewIndex.xlsx"
    if crewindex.exists():
        crewindex_files.append(
            {
                "name": crewindex.name,
                "path": _relative(crewindex),
                "mtime": _to_iso(crewindex.stat().st_mtime),
            }
        )

    names_files = []
    names = MASTER_DIR / "Names.xlsx"
    if names.exists():
        names_files.append(
            {
                "name": names.name,
                "path": _relative(names),
                "mtime": _to_iso(names.stat().st_mtime),
            }
        )

    return {
        "sf_exports": sf_files,
        "combined": combined_files,
        "crewindex": crewindex_files,
        "names": names_files,
    }


@router.post("/browse/query")
def browse_query(payload: BrowseQueryRequest) -> dict:
    file_path = _validate_file_path(payload.file_path)
    df, sheet_name = _read_sheet_cached(file_path, payload.sheet)

    filtered = _apply_filters(df, payload.filters)
    sorted_df = _apply_sort(filtered, payload.sort)
    page_df, total_rows = _paginate(sorted_df, payload.page, payload.page_size)

    rows = page_df.astype(str).to_dict(orient="records")
    return {
        "columns": [str(col) for col in page_df.columns.tolist()],
        "rows": rows,
        "page": payload.page,
        "page_size": payload.page_size,
        "total_rows": total_rows,
        "sheet": sheet_name,
    }


@router.post("/browse/preview")
def browse_preview(payload: BrowsePreviewRequest) -> dict:
    file_path = _validate_file_path(payload.file_path)
    df, sheet_name = _read_sheet_cached(file_path, payload.sheet)
    preview_df = df.head(payload.limit)

    return {
        "columns": [str(col) for col in preview_df.columns.tolist()],
        "rows": preview_df.astype(str).to_dict(orient="records"),
        "total_rows": int(len(df)),
        "sheet": sheet_name,
    }


@router.post("/browse/distinct")
def browse_distinct(payload: BrowseDistinctRequest) -> dict:
    file_path = _validate_file_path(payload.file_path)
    df, sheet_name = _read_sheet_cached(file_path, payload.sheet)
    filtered = _apply_filters(df, payload.filters)

    requested = [c for c in payload.columns if c in filtered.columns]
    if not requested:
        return {"sheet": sheet_name, "columns": [], "rows": []}

    out = (
        filtered[requested]
        .fillna("")
        .astype(str)
        .drop_duplicates()
        .head(payload.limit)
    )
    return {
        "sheet": sheet_name,
        "columns": requested,
        "rows": out.to_dict(orient="records"),
    }


class BrowseExportRequest(BaseModel):
    file_path: str = ""
    sheet: str = ""
    sort: list[SortSpec] = Field(default_factory=list)
    filters: list[FilterSpec] = Field(default_factory=list)
    filename: str = "export"


@router.post("/browse/export")
def browse_export(payload: BrowseExportRequest):
    file_path = _validate_file_path(payload.file_path)
    df, sheet_name = _read_sheet_cached(file_path, payload.sheet)

    filtered = _apply_filters(df, payload.filters)
    sorted_df = _apply_sort(filtered, payload.sort)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        sorted_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    buffer.seek(0)

    safe_name = "".join(c for c in payload.filename if c.isalnum() or c in "._- ")
    filename = f"{safe_name}.xlsx"

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


REGISTRY_EXPORT_COLS = [
    "CM ID", "Actual Name", "Actual Title", "Last General Title",
    "Last Department", "Actual Phone", "Actual Email",
    "Last Phone", "Last Email", "Status", "Shows Worked",
]


def _load_registry_df() -> pd.DataFrame:
    if not CREW_REGISTRY_FILE.exists():
        raise HTTPException(status_code=404, detail="CrewRegistry.xlsx not found")
    try:
        df = pd.read_excel(CREW_REGISTRY_FILE, sheet_name="CrewRegistry",
                           dtype=str, engine="openpyxl")
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Could not read CrewRegistry: {exc}") from exc
    return df.fillna("")


def _filter_registry_by_gcmids(df: pd.DataFrame, gcmids: list[str]) -> pd.DataFrame:
    gcmid_set = {str(g).split(".")[0].strip() for g in gcmids if g}
    df["_cm_id_str"] = df["CM ID"].astype(str).str.split(".").str[0].str.strip()
    filtered = df[df["_cm_id_str"].isin(gcmid_set)].drop(columns=["_cm_id_str"])
    available = [c for c in REGISTRY_EXPORT_COLS if c in filtered.columns]
    return filtered[available].sort_values("Actual Name", na_position="last")


class RegistryLookupRequest(BaseModel):
    gcmids: list[str]


@router.post("/browse/registry_lookup")
def browse_registry_lookup(payload: RegistryLookupRequest) -> dict:
    df = _load_registry_df()
    out = _filter_registry_by_gcmids(df, payload.gcmids)
    return {
        "columns": list(out.columns),
        "rows": out.to_dict(orient="records"),
        "total_rows": len(out),
    }


class RegistryExportRequest(BaseModel):
    gcmids: list[str]
    filename: str = "CrewExport"


@router.post("/browse/registry_export")
def browse_registry_export(payload: RegistryExportRequest):
    df = _load_registry_df()
    out = _filter_registry_by_gcmids(df, payload.gcmids)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="CrewExport")
    buffer.seek(0)

    safe_name = "".join(c for c in payload.filename if c.isalnum() or c in "._- ")
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{safe_name}.xlsx"'},
    )

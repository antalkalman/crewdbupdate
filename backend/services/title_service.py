from __future__ import annotations

from pathlib import Path

import openpyxl
import pandas as pd
from fastapi import HTTPException

from backend.utils.paths import MASTER_DIR, SF_ARCHIVE_DIR, TITLEMAP_FILE


class TitleService:
    def __init__(self) -> None:
        self.helper_path = MASTER_DIR / "Helper.xlsx"
        self.titlemap_path = TITLEMAP_FILE

    @staticmethod
    def _normalize_col(name: str) -> str:
        return "".join(ch for ch in str(name).lower() if ch.isalnum())

    def _resolve_key_columns(self, df: pd.DataFrame) -> tuple[str | None, str | None]:
        normalized = {self._normalize_col(col): col for col in df.columns}

        key_col = None
        for cand in ("titleproject", "titleprojectkey", "key"):
            if cand in normalized:
                key_col = normalized[cand]
                break

        general_col = None
        for cand in ("generaltitle", "general_title"):
            if cand in normalized:
                general_col = normalized[cand]
                break

        return key_col, general_col

    def _resolve_project_and_title_columns(self, df: pd.DataFrame) -> tuple[str | None, str | None]:
        normalized = {self._normalize_col(col): col for col in df.columns}

        title_candidates = [
            "projectjobtitle",
            "jobtitle",
            "title",
            "projecttitle",
        ]
        project_candidates = [
            "project",
            "projectname",
            "production",
            "show",
        ]

        title_col = next((normalized[c] for c in title_candidates if c in normalized), None)
        project_col = next((normalized[c] for c in project_candidates if c in normalized), None)

        if not title_col:
            for col in df.columns:
                norm = self._normalize_col(col)
                if "title" in norm and "general" not in norm:
                    title_col = col
                    break

        if not project_col:
            for col in df.columns:
                norm = self._normalize_col(col)
                if "project" in norm and "department" not in norm:
                    project_col = col
                    break

        return project_col, title_col

    def get_latest_sflist_path(self) -> Path:
        pattern = str(SF_ARCHIVE_DIR / "SFlist_*.xlsx")
        candidates = list(SF_ARCHIVE_DIR.glob("SFlist_*.xlsx"))
        if not candidates:
            raise HTTPException(
                status_code=404,
                detail=f"No SF list export found in {SF_ARCHIVE_DIR} with pattern SFlist_*.xlsx",
            )

        latest = max(candidates, key=lambda p: p.stat().st_mtime)
        print(f"[TitleService] SF list discovery pattern: {pattern}")
        print(f"[TitleService] Using latest SF list: {latest}")
        return latest

    def read_helper_title_conv(self, source_path: Path) -> tuple[set[str], pd.DataFrame]:
        if not source_path.exists():
            raise HTTPException(status_code=404, detail=f"Title file not found: {source_path}")

        try:
            workbook = pd.ExcelFile(source_path, engine="openpyxl")
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Could not open title file: {exc}") from exc

        if "Title conv" not in workbook.sheet_names:
            raise HTTPException(status_code=500, detail=f"Missing required sheet 'Title conv' in {source_path.name}")

        df = pd.read_excel(source_path, sheet_name="Title conv", dtype=str, engine="openpyxl")
        print(f"[TitleService] Loaded '{source_path.name}' Title conv with columns: {df.columns.tolist()}")

        key_col, _ = self._resolve_key_columns(df)
        if key_col:
            keys = (
                df[key_col]
                .fillna("")
                .astype(str)
                .str.strip()
            )
            key_set = {k for k in keys if k}
            print(f"[TitleService] Using existing key column: {key_col}")
            return key_set, df

        project_col, title_col = self._resolve_project_and_title_columns(df)
        if not project_col or not title_col:
            raise HTTPException(
                status_code=500,
                detail="Could not determine key columns in 'Title conv'. "
                "Need either a key column or both Project and Project job title columns.",
            )

        keys = (
            df[title_col].fillna("").astype(str).str.strip()
            + "--"
            + df[project_col].fillna("").astype(str).str.strip()
        )
        key_set = {k for k in keys if k != "--"}
        print(
            f"[TitleService] Built keys from columns: project='{project_col}', title='{title_col}'"
        )
        return key_set, df

    def read_valid_general_titles(self, source_path: Path) -> list[str]:
        if not source_path.exists():
            raise HTTPException(status_code=404, detail=f"Title file not found: {source_path}")

        try:
            workbook = pd.ExcelFile(source_path, engine="openpyxl")
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Could not open title file: {exc}") from exc

        if "General Title" not in workbook.sheet_names:
            raise HTTPException(status_code=500, detail=f"Missing required sheet 'General Title' in {source_path.name}")

        df = pd.read_excel(source_path, sheet_name="General Title", dtype=str, engine="openpyxl")
        project_col, title_col = self._resolve_project_and_title_columns(df)
        _, general_col = self._resolve_key_columns(df)

        if general_col:
            source_col = general_col
        elif "Title" in df.columns:
            source_col = "Title"
        elif title_col:
            source_col = title_col
        else:
            source_col = df.columns[0]

        values = (
            df[source_col]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        valid = sorted({v for v in values if v})
        print(f"[TitleService] Loaded {len(valid)} valid general titles from column '{source_col}'")
        if project_col:
            print(f"[TitleService] General Title sheet project-like column detected: {project_col}")
        return valid

    def compute_unmapped_title_pairs(self, sflist_path: Path, helper_keys: set[str]) -> list[dict]:
        if not sflist_path.exists():
            raise HTTPException(status_code=404, detail=f"SF list file not found: {sflist_path}")

        df = pd.read_excel(sflist_path, dtype=str, engine="openpyxl")
        project_col, title_col = self._resolve_project_and_title_columns(df)
        print(f"[TitleService] SF list columns detected: {df.columns.tolist()}")
        print(f"[TitleService] SF list selected columns: project='{project_col}', title='{title_col}'")

        if not project_col or not title_col:
            raise HTTPException(
                status_code=500,
                detail="Could not detect required columns in SF list for Project and Project job title.",
            )

        # Detect department column
        dept_col = None
        for cand in ("projectdepartment", "department", "dept"):
            if cand in {self._normalize_col(c): c for c in df.columns}:
                dept_col = {self._normalize_col(c): c for c in df.columns}[cand]
                break

        out = pd.DataFrame(
            {
                "project": df[project_col].fillna("").astype(str).str.strip(),
                "project_job_title": df[title_col].fillna("").astype(str).str.strip(),
                "department": df[dept_col].fillna("").astype(str).str.strip() if dept_col else "",
            }
        )
        out = out[(out["project"] != "") & (out["project_job_title"] != "")]
        out["title_project_key"] = out["project_job_title"] + "--" + out["project"]
        out = out.drop_duplicates(subset=["title_project_key"]).sort_values(
            by=["project", "project_job_title"]
        )
        out["general_title"] = ""

        if helper_keys:
            out = out[~out["title_project_key"].isin(helper_keys)]

        return out[["project", "project_job_title", "department", "title_project_key", "general_title"]].to_dict(
            orient="records"
        )

    def _write_to_titlemap(self, validated_pairs: list[tuple[str, str]]) -> dict[str, int]:
        """Append validated (key, general_title) pairs to TitleMap.xlsx."""
        if not self.titlemap_path.exists():
            print(f"[TitleService] WARNING: TitleMap.xlsx not found at {self.titlemap_path}, skipping")
            return {"titlemap_added": 0, "titlemap_skipped": 0}

        try:
            wb = openpyxl.load_workbook(self.titlemap_path)
        except PermissionError:
            print("[TitleService] WARNING: TitleMap.xlsx is open, skipping TitleMap write")
            return {"titlemap_added": 0, "titlemap_skipped": 0}

        ws = wb["Title conv"]

        # Read existing keys from col A
        existing = set()
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
            val = row[0].value
            if val is not None:
                existing.add(str(val).strip())

        added = 0
        skipped = 0
        for key, gt in validated_pairs:
            if key in existing:
                skipped += 1
                continue
            ws.append([key, gt])
            existing.add(key)
            added += 1

        try:
            wb.save(self.titlemap_path)
            print(f"[TitleService] TitleMap.xlsx: {added} added, {skipped} skipped")
        except PermissionError:
            print("[TitleService] WARNING: Could not save TitleMap.xlsx")
            return {"titlemap_added": 0, "titlemap_skipped": 0}

        return {"titlemap_added": added, "titlemap_skipped": skipped}

    def append_title_mappings_to_helper(self, helper_path: Path, rows: list[dict]) -> dict[str, int]:
        valid_titles = set(self.read_valid_general_titles(self.titlemap_path))
        existing_keys, helper_df = self.read_helper_title_conv(helper_path)
        key_col, general_col = self._resolve_key_columns(helper_df)
        project_col, title_col = self._resolve_project_and_title_columns(helper_df)

        if not key_col and (not project_col or not title_col):
            raise HTTPException(
                status_code=500,
                detail="Cannot append mappings because helper sheet column structure is unclear.",
            )

        if not general_col:
            general_col = "General Title"
            if general_col not in helper_df.columns:
                helper_df[general_col] = pd.NA

        appended_rows: list[dict] = []
        validated_pairs: list[tuple[str, str]] = []
        skipped_invalid = 0
        skipped_duplicates = 0

        for row in rows:
            key = str(row.get("title_project_key", "")).strip()
            general_title = str(row.get("general_title", "")).strip()
            project = str(row.get("project", "")).strip()
            project_job_title = str(row.get("project_job_title", "")).strip()

            # Defensive: rebuild malformed or missing key from explicit fields.
            if (not key or "--" not in key) and project_job_title and project:
                key = f"{project_job_title}--{project}"

            if not key or not general_title or general_title not in valid_titles:
                skipped_invalid += 1
                continue

            if key in existing_keys:
                skipped_duplicates += 1
                continue

            if (not project or not project_job_title) and "--" in key:
                parsed_title, parsed_project = key.split("--", 1)
                project_job_title = project_job_title or parsed_title.strip()
                project = project or parsed_project.strip()

            new_row = {col: pd.NA for col in helper_df.columns}
            new_row[general_col] = general_title
            if key_col:
                new_row[key_col] = key
            if project_col and project_col != key_col:
                new_row[project_col] = project
            if title_col and title_col != key_col:
                new_row[title_col] = project_job_title

            appended_rows.append(new_row)
            validated_pairs.append((key, general_title))
            existing_keys.add(key)

        # Write to Helper.xlsx
        helper_written = len(appended_rows)
        if appended_rows:
            new_df = pd.DataFrame(appended_rows)
            helper_df = pd.concat([helper_df, new_df], ignore_index=True)

            try:
                with pd.ExcelWriter(
                    helper_path,
                    engine="openpyxl",
                    mode="a",
                    if_sheet_exists="replace",
                ) as writer:
                    helper_df.to_excel(writer, sheet_name="Title conv", index=False)
            except PermissionError as exc:
                print(f"[TitleService] WARNING: Could not write Helper.xlsx: {exc}")
                helper_written = 0
            except OSError as exc:
                print(f"[TitleService] WARNING: Could not write Helper.xlsx: {exc}")
                helper_written = 0

        # Write to TitleMap.xlsx
        titlemap_result = {"titlemap_added": 0, "titlemap_skipped": 0}
        if validated_pairs:
            titlemap_result = self._write_to_titlemap(validated_pairs)

        return {
            "appended": helper_written,
            "skipped_invalid": skipped_invalid,
            "skipped_duplicates": skipped_duplicates,
            **titlemap_result,
        }

    def get_unmapped_titles(self) -> dict:
        sflist_path = self.get_latest_sflist_path()
        # Read from TitleMap.xlsx (primary source)
        helper_keys, _ = self.read_helper_title_conv(self.titlemap_path)
        rows = self.compute_unmapped_title_pairs(sflist_path, helper_keys)
        valid_general_titles = self.read_valid_general_titles(self.titlemap_path)
        return {"rows": rows, "valid_general_titles": valid_general_titles}

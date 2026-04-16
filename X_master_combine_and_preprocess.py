from __future__ import annotations

import glob
import subprocess
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd


def _validate_inputs(base_path: Path) -> None:
    master_dir = base_path / "Master_database"
    sf_archive_dir = base_path / "SF_Archive"

    if not master_dir.is_dir():
        raise FileNotFoundError(f"Missing folder: {master_dir}")
    if not sf_archive_dir.is_dir():
        raise FileNotFoundError(f"Missing folder: {sf_archive_dir}")

    required_files = [
        master_dir / "Names.xlsx",
        master_dir / "Helper.xlsx",
        master_dir / "combined_field_mapping.xlsx",
    ]
    missing = [str(path) for path in required_files if not path.exists()]
    if missing:
        raise FileNotFoundError("Missing required file(s): " + ", ".join(missing))

    if not glob.glob(str(master_dir / "Historical_data_*.xlsx")):
        raise FileNotFoundError(
            f"No historical input found with pattern: {master_dir / 'Historical_data_*.xlsx'}"
        )

    if not glob.glob(str(sf_archive_dir / "SFlist_*.xlsx")):
        raise FileNotFoundError(
            f"No SF list export found with pattern: {sf_archive_dir / 'SFlist_*.xlsx'}"
        )


def _run_script(script_path: Path) -> None:
    try:
        result = subprocess.run(
            ["python3", str(script_path)],
            check=True,
            capture_output=True,
            text=True,
        )
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr)
    except subprocess.CalledProcessError as exc:
        combined_output = f"{exc.stdout or ''}\n{exc.stderr or ''}".lower()
        if "permissionerror" in combined_output or "used by another process" in combined_output:
            raise PermissionError(
                "Could not write Combined_All_CrewData.xlsx. Please close the Excel file and retry."
            ) from exc
        raise RuntimeError(f"Script failed: {script_path.name}\n{exc.stdout or ''}\n{exc.stderr or ''}") from exc


def run_master_combine_and_preprocess(base_path: Path | str | None = None) -> dict:
    root = Path(base_path) if base_path else Path(__file__).resolve().parent
    combined_script = root / "x_master_combined.py"
    preprocess_script = root / "x_master_preprocess.py"
    output_path = root / "Master_database" / "Combined_All_CrewData.xlsx"

    if not combined_script.exists():
        raise FileNotFoundError(f"Missing script: {combined_script}")
    if not preprocess_script.exists():
        raise FileNotFoundError(f"Missing script: {preprocess_script}")

    _validate_inputs(root)

    print("[master] Running x_master_combined.py ...")
    _run_script(combined_script)
    print("[master] Running x_master_preprocess.py ...")
    _run_script(preprocess_script)

    if not output_path.exists():
        raise FileNotFoundError(f"Expected output not found: {output_path}")

    try:
        df_combined = pd.read_excel(output_path, sheet_name="Combined", dtype=str, engine="openpyxl")
        rows_written = int(len(df_combined))
    except PermissionError as exc:
        raise PermissionError(
            "Combined_All_CrewData.xlsx appears to be open. Close Excel and retry."
        ) from exc

    return {
        "ok": True,
        "output_file": str(output_path),
        "created_at": datetime.now(timezone.utc).isoformat(),
        "meta": {
            "rows_written": rows_written,
        },
    }


if __name__ == "__main__":
    summary = run_master_combine_and_preprocess()
    print(summary)

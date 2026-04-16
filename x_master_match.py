from __future__ import annotations

import os
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

import Levenshtein
import pandas as pd

NICKNAME_MAP = {
    "gabi": "gabriella",
    "zsuzsa": "zsuzsanna",
    "zsuzsi": "zsuzsanna",
    "gergo": "gergely",
    "kati": "katalin",
    "erzsi": "erzsebet",
    "bobe": "erzsebet",
    "bori": "borbala",
    "dani": "daniel",
    "moni": "monika",
    "zoli": "zoltan",
    "niki": "nikoletta",
    "pisti": "istvan",
    "magdi": "magdolna",
    "jr": "junior",
    "jrxx": "junior",
    "orsi": "orsolya",
    "ricsi": "richard",
    "gyuri": "gyorgy",
}


def strip_accents(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", str(text)) if unicodedata.category(c) != "Mn")


def normalize_phone(phone: str) -> str:
    raw = re.sub(r"\D", "", str(phone or ""))
    if raw.startswith("36"):
        raw = raw[2:]
    elif raw.startswith("06"):
        raw = raw[2:]
    elif raw.startswith("6"):
        raw = raw[1:]
    return f"36{raw}" if raw else ""


def clean_text(text: str) -> str:
    normalized = strip_accents(str(text or "")).lower().strip()
    return re.sub(r"[\"'’().\s]", "", normalized)


def tokenize_name(name: str) -> list[str]:
    if not isinstance(name, str):
        return []
    normalized = strip_accents(name.lower())
    normalized = re.sub(r"[\"'’().]", "", normalized)
    tokens = re.findall(r"\b\w+\b", normalized)
    return [NICKNAME_MAP.get(tok, tok) for tok in tokens if tok != "né"]


def token_match_score(input_token: str, target_token: str) -> float:
    if input_token == target_token:
        return 1.0
    if target_token.startswith(input_token) and len(input_token) >= 2:
        return 0.75
    dist = Levenshtein.distance(input_token, target_token)
    if dist == 1:
        return 0.5
    if dist == 2:
        return 0.25
    return 0.0


def _normalize_gcmid(value) -> str:
    try:
        if pd.isna(value):
            return ""
        return str(int(float(value)))
    except Exception:
        return str(value or "").strip().replace(".0", "")


def _load_inputs(base_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    input_path = base_dir / "Combined_All_CrewData.xlsx"
    names_path = base_dir / "Names.xlsx"

    df_input = pd.read_excel(input_path, sheet_name="Combined", dtype=str, engine="openpyxl")

    df_tokens = pd.read_excel(input_path, sheet_name="Tokenized Names", dtype=str, engine="openpyxl")
    df_tokens = df_tokens.rename(columns={"GCMID": "CM ID"})

    df_emails = pd.read_excel(input_path, sheet_name="Emails", dtype=str, engine="openpyxl")
    df_emails = df_emails.rename(columns={"GCMID": "CM ID"})

    df_phones = pd.read_excel(input_path, sheet_name="Phones", dtype=str, engine="openpyxl")
    df_phones = df_phones.rename(columns={"GCMID": "CM ID"})

    df_dept = pd.read_excel(input_path, sheet_name="General Departments", dtype=str, engine="openpyxl")
    df_dept = df_dept[["GCMID", "General Department"]].copy()
    df_dept = df_dept.rename(columns={"General Department": "DB General Department"})
    df_dept["GCMID"] = df_dept["GCMID"].apply(_normalize_gcmid)

    df_names_info = pd.read_excel(names_path, sheet_name="Names", dtype=str, engine="openpyxl")
    df_names_info = df_names_info[["CM ID", "Sure Name", "First Name", "Actual Title"]].copy()
    df_names_info = df_names_info.rename(
        columns={
            "CM ID": "GCMID",
            "Sure Name": "DB Surname",
            "First Name": "DB Firstname",
            "Actual Title": "DB Title",
        }
    )
    df_names_info["GCMID"] = df_names_info["GCMID"].apply(_normalize_gcmid)
    df_names_info["Matched Name"] = (
        df_names_info["DB Surname"].fillna("").astype(str).str.strip()
        + " "
        + df_names_info["DB Firstname"].fillna("").astype(str).str.strip()
    ).str.strip()

    return df_input, df_tokens, df_emails, df_phones, df_dept, df_names_info


def _build_lookup_maps(
    df_tokens: pd.DataFrame,
    df_emails: pd.DataFrame,
    df_phones: pd.DataFrame,
    df_dept: pd.DataFrame,
    df_names_info: pd.DataFrame,
) -> tuple[dict[str, list[str]], dict[str, set[str]], dict[str, set[str]], dict[str, str], dict[str, dict]]:
    token_map: dict[str, list[str]] = {}
    for _, row in df_tokens.iterrows():
        gcmid = _normalize_gcmid(row.get("CM ID"))
        token = clean_text(row.get("Token", ""))
        if not gcmid or not token:
            continue
        token_map.setdefault(gcmid, []).append(token)

    email_map: dict[str, set[str]] = {}
    for _, row in df_emails.iterrows():
        gcmid = _normalize_gcmid(row.get("CM ID"))
        email = clean_text(row.get("Email", ""))
        if not gcmid or not email:
            continue
        email_map.setdefault(gcmid, set()).add(email)

    phone_map: dict[str, set[str]] = {}
    for _, row in df_phones.iterrows():
        gcmid = _normalize_gcmid(row.get("CM ID"))
        phone = normalize_phone(row.get("Phone", ""))
        if not gcmid or not phone:
            continue
        phone_map.setdefault(gcmid, set()).add(phone)

    dept_map = {
        _normalize_gcmid(row.get("GCMID")): str(row.get("DB General Department", "") or "").strip()
        for _, row in df_dept.iterrows()
        if _normalize_gcmid(row.get("GCMID"))
    }

    names_map = {
        _normalize_gcmid(row.get("GCMID")): {
            "matched_name": str(row.get("Matched Name", "") or "").strip(),
            "db_surname": str(row.get("DB Surname", "") or "").strip(),
            "db_firstname": str(row.get("DB Firstname", "") or "").strip(),
            "db_title": str(row.get("DB Title", "") or "").strip(),
        }
        for _, row in df_names_info.iterrows()
        if _normalize_gcmid(row.get("GCMID"))
    }

    return token_map, email_map, phone_map, dept_map, names_map


def _score_candidates(
    source_name: str,
    source_email: str,
    source_phone: str,
    source_general_dept: str,
    token_map: dict[str, list[str]],
    email_map: dict[str, set[str]],
    phone_map: dict[str, set[str]],
    dept_map: dict[str, str],
) -> list[dict]:
    input_tokens = tokenize_name(source_name)
    input_email = clean_text(source_email)
    input_phone = normalize_phone(source_phone)
    input_dept = str(source_general_dept or "").strip()

    all_ids = set(token_map.keys()) | set(email_map.keys()) | set(phone_map.keys())
    candidates: list[dict] = []

    for gcmid in all_ids:
        tokens = token_map.get(gcmid, [])
        name_score = 0.0
        if input_tokens and tokens:
            name_score = sum(max((token_match_score(t, tok) for tok in tokens), default=0.0) for t in input_tokens)

        email_score = 0.0
        if input_email:
            for target_email in email_map.get(gcmid, set()):
                if target_email == input_email:
                    email_score = max(email_score, 1.0)
                elif Levenshtein.distance(target_email, input_email) == 1:
                    email_score = max(email_score, 0.5)

        phone_score = 0.0
        if input_phone:
            for target_phone in phone_map.get(gcmid, set()):
                if target_phone == input_phone:
                    phone_score = max(phone_score, 1.0)
                elif Levenshtein.distance(target_phone, input_phone) == 1:
                    phone_score = max(phone_score, 0.5)

        dept_score = 0.0
        if input_dept and dept_map.get(gcmid, "") == input_dept:
            dept_score = 0.5

        final_score = name_score * 1.5 + email_score + phone_score + dept_score
        if final_score <= 0:
            continue

        candidates.append(
            {
                "suggested_gcmid": gcmid,
                "name_score": round(float(name_score), 4),
                "email_score": round(float(email_score), 4),
                "phone_score": round(float(phone_score), 4),
                "dept_score": round(float(dept_score), 4),
                "final_score": round(float(final_score), 4),
                "db_general_department": dept_map.get(gcmid, ""),
            }
        )

    candidates.sort(key=lambda x: (-x["final_score"], -x["name_score"], x["suggested_gcmid"]))
    return candidates


def run_matching(base_dir: str | Path | None = None, write_output: bool = False) -> dict:
    root = Path(base_dir) if base_dir else Path(__file__).resolve().parent / "Master_database"
    output_path = root / "Matched_GCMID_from_helpers.xlsx"

    print(f"[match] Base dir: {root}")
    df_input, df_tokens, df_emails, df_phones, df_dept, df_names_info = _load_inputs(root)

    token_map, email_map, phone_map, dept_map, names_map = _build_lookup_maps(
        df_tokens, df_emails, df_phones, df_dept, df_names_info
    )

    df_missing = df_input[
        df_input["Crew list name"].notna()
        & (
            df_input["GCMID"].isna()
            | (df_input["GCMID"].astype(str).str.strip() == "")
        )
    ].copy()

    confirmed: list[dict] = []
    possible: list[dict] = []
    new_names: list[dict] = []

    for row_idx, row in df_missing.iterrows():
        source = {
            "source_row_index": int(row_idx),
            "source_key": f"{str(row.get('Crew member id', '')).strip()}--{str(row.get('Project', '')).strip()}",
            "crew_member_id": str(row.get("Crew member id", "") or "").strip(),
            "project": str(row.get("Project", "") or "").strip(),
            "project_job_title": str(row.get("Project job title", "") or "").strip(),
            "name_on_crew_list": str(row.get("Crew list name", "") or "").strip(),
            "general_department": str(row.get("General Department", "") or "").strip(),
            "general_title": str(row.get("General Title", "") or "").strip(),
            "email": str(row.get("Crew email", "") or "").strip(),
            "phone": str(row.get("Mobile number", "") or "").strip(),
        }

        candidates = _score_candidates(
            source_name=source["name_on_crew_list"],
            source_email=source["email"],
            source_phone=source["phone"],
            source_general_dept=source["general_department"],
            token_map=token_map,
            email_map=email_map,
            phone_map=phone_map,
            dept_map=dept_map,
        )

        for cand in candidates:
            info = names_map.get(cand["suggested_gcmid"], {})
            cand["matched_name"] = info.get("matched_name", "")
            cand["db_title"] = info.get("db_title", "")
            cand["db_surname"] = info.get("db_surname", "")
            cand["db_firstname"] = info.get("db_firstname", "")

        best = candidates[0] if candidates else None

        if best and best["name_score"] >= 1.25 and (best["email_score"] + best["phone_score"]) >= 1.0:
            confirmed.append(
                {
                    **source,
                    "suggested_gcmid": best["suggested_gcmid"],
                    "matched_name": best.get("matched_name", ""),
                    "db_title": best.get("db_title", ""),
                    "name_score": best["name_score"],
                    "email_score": best["email_score"],
                    "phone_score": best["phone_score"],
                    "dept_score": best["dept_score"],
                    "final_score": best["final_score"],
                }
            )
            continue

        top_candidates = []
        seen = set()
        for cand in candidates:
            gcmid = cand["suggested_gcmid"]
            if gcmid in seen:
                continue
            if cand["final_score"] < 1.25:
                continue
            seen.add(gcmid)
            top_candidates.append(cand)
            if len(top_candidates) >= 5:
                break

        if top_candidates:
            possible.append(
                {
                    **source,
                    "candidate_count": len(top_candidates),
                    "candidates": top_candidates,
                }
            )
            if source["email"] and source["phone"]:
                new_names.append(
                    {
                        **source,
                        "reason": f"Possible match exists ({len(top_candidates)} candidate(s)) but has phone & email",
                    }
                )
        else:
            reason = "No candidates scored above threshold"
            if best and best["final_score"] > 0:
                reason = f"Top score too low ({best['final_score']})"

            new_names.append(
                {
                    **source,
                    "reason": reason,
                }
            )

    meta = {
        "created_at": datetime.now(timezone.utc).isoformat(),
        "missing_count": int(len(df_missing)),
        "confirmed_count": int(len(confirmed)),
        "possible_count": int(len(possible)),
        "new_names_count": int(len(new_names)),
    }

    if write_output:
        confirmed_df = pd.DataFrame(confirmed)

        flattened_possible = []
        for row in possible:
            base = {k: v for k, v in row.items() if k != "candidates"}
            for rank, cand in enumerate(row.get("candidates", []), start=1):
                flattened_possible.append({**base, "rank": rank, **cand})
        possible_df = pd.DataFrame(flattened_possible)

        new_names_df = pd.DataFrame(new_names)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            confirmed_df.to_excel(writer, sheet_name="Confirmed Matches", index=False)
            possible_df.to_excel(writer, sheet_name="Possible Matches", index=False)
            new_names_df.to_excel(writer, sheet_name="Possible New Names", index=False)

        print(f"[match] Output written: {output_path}")

    return {
        "confirmed": confirmed,
        "possible": possible,
        "new_names": new_names,
        "meta": meta,
    }


if __name__ == "__main__":
    result = run_matching(write_output=True)
    print("[match] Completed", result["meta"])

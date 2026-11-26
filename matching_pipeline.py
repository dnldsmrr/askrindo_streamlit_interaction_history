from __future__ import annotations

import io
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.styles import Font

from pipeline_etl_askrindo import (
    company_match_key,
    normalize_company_name,
    normalize_email_detailed,
    normalize_phone_detailed,
    sanitize_filename,
)

MATCHING_VERSION = "1.0.0"


@dataclass
class MatchingConfig:
    reference_sheet: str | int
    reference_header: int
    reference_row_id_column: str
    reference_company_column: str
    reference_phone_columns: list[str] = field(default_factory=list)
    reference_email_columns: list[str] = field(default_factory=list)

    source_sheet: str | int = 0
    source_header: int = 0
    source_phone_columns: list[str] = field(default_factory=list)
    source_email_columns: list[str] = field(default_factory=list)
    source_company_column: str | None = None


def read_excel_bytes(file_bytes: bytes, sheet_name: str | int, header: int) -> pd.DataFrame:
    with io.BytesIO(file_bytes) as buffer:
        return pd.read_excel(buffer, sheet_name=sheet_name, header=header)


def get_sheet_names(file_bytes: bytes) -> list[str]:
    with io.BytesIO(file_bytes) as buffer:
        return pd.ExcelFile(buffer).sheet_names


def _split_semicolon(value: Any) -> list[str]:
    if pd.isna(value):
        return []
    return [item.strip() for item in str(value).split(";") if item.strip()]


def _normalize_email_list(value: Any) -> list[str]:
    result = normalize_email_detailed(value)
    return _split_semicolon(result.normalized)


def _normalize_phone_list(value: Any) -> list[str]:
    result = normalize_phone_detailed(value)
    return _split_semicolon(result.normalized)


def validate_matching_config(
    config: MatchingConfig,
    reference_columns: list[str],
    source_columns: list[str],
) -> list[str]:
    errors: list[str] = []
    ref = set(reference_columns)
    src = set(source_columns)

    for column in [config.reference_row_id_column, config.reference_company_column]:
        if column not in ref:
            errors.append(f"Kolom database tidak ditemukan: {column}")
    for column in config.reference_phone_columns + config.reference_email_columns:
        if column not in ref:
            errors.append(f"Kolom database tidak ditemukan: {column}")
    for column in config.source_phone_columns + config.source_email_columns:
        if column not in src:
            errors.append(f"Kolom file customer tidak ditemukan: {column}")
    if config.source_company_column and config.source_company_column not in src:
        errors.append(f"Kolom nama perusahaan sumber tidak ditemukan: {config.source_company_column}")
    if not config.reference_phone_columns and not config.reference_email_columns:
        errors.append("Database harus memiliki minimal satu kolom telepon atau email.")
    if not config.source_phone_columns and not config.source_email_columns:
        errors.append("File customer harus memiliki minimal satu kolom telepon atau email.")
    return errors


def prepare_reference_database(reference_df: pd.DataFrame, config: MatchingConfig) -> pd.DataFrame:
    work = reference_df.copy().reset_index(drop=True)
    row_ids = work[config.reference_row_id_column].astype("string").str.strip()
    if row_ids.isna().any() or row_ids.eq("").any():
        raise ValueError("Database memiliki ROW_ID kosong.")
    duplicates = row_ids[row_ids.duplicated(keep=False)]
    if not duplicates.empty:
        sample = ", ".join(duplicates.drop_duplicates().head(10).tolist())
        raise ValueError(f"ROW_ID database harus unik. Duplikat ditemukan: {sample}")

    work["REFERENCE_ROW_ID"] = row_ids
    work["REFERENCE_COMPANY_ORIGINAL"] = work[config.reference_company_column]
    work["REFERENCE_COMPANY_NORMALIZED"] = work[config.reference_company_column].map(normalize_company_name)
    work["REFERENCE_COMPANY_MATCH_KEY"] = work[config.reference_company_column].map(company_match_key)

    phone_values: list[str] = []
    email_values: list[str] = []
    for _, row in work.iterrows():
        phones: list[str] = []
        for column in config.reference_phone_columns:
            phones.extend(_normalize_phone_list(row.get(column)))
        emails: list[str] = []
        for column in config.reference_email_columns:
            emails.extend(_normalize_email_list(row.get(column)))
        phone_values.append("; ".join(dict.fromkeys(phones)))
        email_values.append("; ".join(dict.fromkeys(emails)))
    work["REFERENCE_PHONES_NORMALIZED"] = phone_values
    work["REFERENCE_EMAILS_NORMALIZED"] = email_values
    return work


def build_reference_indexes(reference_prepared: pd.DataFrame) -> dict[str, Any]:
    phone_index: dict[str, set[str]] = defaultdict(set)
    email_index: dict[str, set[str]] = defaultdict(set)
    row_lookup: dict[str, dict[str, Any]] = {}

    for _, row in reference_prepared.iterrows():
        row_id = str(row["REFERENCE_ROW_ID"])
        row_lookup[row_id] = {
            "company_original": row["REFERENCE_COMPANY_ORIGINAL"],
            "company_normalized": row["REFERENCE_COMPANY_NORMALIZED"],
            "company_key": row["REFERENCE_COMPANY_MATCH_KEY"],
        }
        for phone in _split_semicolon(row["REFERENCE_PHONES_NORMALIZED"]):
            phone_index[phone].add(row_id)
        for email in _split_semicolon(row["REFERENCE_EMAILS_NORMALIZED"]):
            email_index[email].add(row_id)

    return {"phone": phone_index, "email": email_index, "rows": row_lookup}


def _source_contact_values(row: pd.Series, config: MatchingConfig) -> tuple[list[str], list[str]]:
    phones: list[str] = []
    emails: list[str] = []
    for column in config.source_phone_columns:
        phones.extend(_normalize_phone_list(row.get(column)))
    for column in config.source_email_columns:
        emails.extend(_normalize_email_list(row.get(column)))
    return list(dict.fromkeys(phones)), list(dict.fromkeys(emails))


def match_one_record(
    phones: list[str],
    emails: list[str],
    company_value: Any,
    indexes: dict[str, Any],
) -> dict[str, Any]:
    phone_candidates: set[str] = set()
    email_candidates: set[str] = set()
    for value in phones:
        phone_candidates.update(indexes["phone"].get(value, set()))
    for value in emails:
        email_candidates.update(indexes["email"].get(value, set()))

    intersection = phone_candidates & email_candidates
    company_key = company_match_key(company_value)

    method = "NONE"
    candidates: set[str] = set()
    score = 0

    if intersection:
        candidates = intersection
        method = "PHONE_EMAIL"
        score = 100
    elif phone_candidates:
        candidates = phone_candidates
        method = "PHONE"
        score = 80
    elif email_candidates:
        candidates = email_candidates
        method = "EMAIL"
        score = 75

    original_candidate_count = len(candidates)
    company_supported = False
    if len(candidates) > 1 and company_key:
        narrowed = {
            row_id for row_id in candidates
            if indexes["rows"][row_id]["company_key"] == company_key
        }
        if narrowed:
            candidates = narrowed
            company_supported = True
            method = f"{method}_COMPANY"
            score = 90 if method.startswith("PHONE") else 85

    if len(candidates) == 1:
        row_id = next(iter(candidates))
        info = indexes["rows"][row_id]
        reason = {
            "PHONE_EMAIL": "Nomor telepon dan email cocok pada satu ROW_ID database.",
            "PHONE": "Nomor telepon cocok pada satu ROW_ID database.",
            "EMAIL": "Email cocok pada satu ROW_ID database.",
        }.get(method, "Nama perusahaan mempersempit beberapa kandidat menjadi satu ROW_ID.")
        return {
            "MATCHED_ROW_ID": row_id,
            "MATCHED_COMPANY_NAME": info["company_original"],
            "MATCHED_COMPANY_NORMALIZED": info["company_normalized"],
            "MATCH_STATUS": "MATCHED",
            "MATCH_METHOD": method,
            "MATCH_SCORE": score,
            "MATCH_CANDIDATE_COUNT": original_candidate_count or 1,
            "MATCH_CANDIDATE_ROW_IDS": row_id,
            "MATCH_REASON": reason,
            "COMPANY_USED_AS_SUPPORT": company_supported,
        }

    if len(candidates) > 1:
        return {
            "MATCHED_ROW_ID": pd.NA,
            "MATCHED_COMPANY_NAME": pd.NA,
            "MATCHED_COMPANY_NORMALIZED": pd.NA,
            "MATCH_STATUS": "AMBIGUOUS",
            "MATCH_METHOD": method,
            "MATCH_SCORE": 50,
            "MATCH_CANDIDATE_COUNT": len(candidates),
            "MATCH_CANDIDATE_ROW_IDS": " | ".join(sorted(candidates)),
            "MATCH_REASON": "Kontak cocok ke lebih dari satu ROW_ID database dan tidak dapat dipersempit secara unik.",
            "COMPANY_USED_AS_SUPPORT": company_supported,
        }

    return {
        "MATCHED_ROW_ID": pd.NA,
        "MATCHED_COMPANY_NAME": pd.NA,
        "MATCHED_COMPANY_NORMALIZED": pd.NA,
        "MATCH_STATUS": "NOT_FOUND",
        "MATCH_METHOD": "NONE",
        "MATCH_SCORE": 0,
        "MATCH_CANDIDATE_COUNT": 0,
        "MATCH_CANDIDATE_ROW_IDS": "",
        "MATCH_REASON": "Nomor telepon dan email tidak ditemukan pada database referensi.",
        "COMPANY_USED_AS_SUPPORT": False,
    }


def run_matching(
    reference_df: pd.DataFrame,
    source_df: pd.DataFrame,
    config: MatchingConfig,
) -> dict[str, pd.DataFrame]:
    errors = validate_matching_config(config, list(reference_df.columns), list(source_df.columns))
    if errors:
        raise ValueError(" | ".join(errors))

    reference_prepared = prepare_reference_database(reference_df, config)
    indexes = build_reference_indexes(reference_prepared)

    source = source_df.copy().reset_index(drop=True)
    source.insert(0, "SOURCE_ROW_NUMBER", range(config.source_header + 2, config.source_header + 2 + len(source)))
    source.insert(0, "SOURCE_ROW_ID", [f"SRC-{number:07d}" for number in range(1, len(source) + 1)])

    source_company_original: list[Any] = []
    source_company_normalized: list[Any] = []
    source_company_key: list[str] = []
    source_phones: list[str] = []
    source_emails: list[str] = []
    match_rows: list[dict[str, Any]] = []

    for _, row in source.iterrows():
        phones, emails = _source_contact_values(row, config)
        company = row.get(config.source_company_column) if config.source_company_column else pd.NA
        source_company_original.append(company)
        source_company_normalized.append(normalize_company_name(company))
        source_company_key.append(company_match_key(company))
        source_phones.append("; ".join(phones))
        source_emails.append("; ".join(emails))
        match_rows.append(match_one_record(phones, emails, company, indexes))

    source["SOURCE_PHONE_NORMALIZED"] = source_phones
    source["SOURCE_EMAIL_NORMALIZED"] = source_emails
    source["SOURCE_COMPANY_ORIGINAL"] = source_company_original
    source["SOURCE_COMPANY_NORMALIZED"] = source_company_normalized
    source["SOURCE_COMPANY_MATCH_KEY"] = source_company_key
    result = pd.concat([source, pd.DataFrame(match_rows)], axis=1)

    matched = result[result["MATCH_STATUS"] == "MATCHED"].copy()
    ambiguous = result[result["MATCH_STATUS"] == "AMBIGUOUS"].copy()
    not_found = result[result["MATCH_STATUS"] == "NOT_FOUND"].copy()

    summary_rows = [
        {"METRIC": "Source Rows", "VALUE": len(result)},
        {"METRIC": "Matched Rows", "VALUE": len(matched)},
        {"METRIC": "Ambiguous Rows", "VALUE": len(ambiguous)},
        {"METRIC": "Not Found Rows", "VALUE": len(not_found)},
    ]
    for method, count in result["MATCH_METHOD"].value_counts(dropna=False).items():
        summary_rows.append({"METRIC": f"Method: {method}", "VALUE": int(count)})

    config_df = pd.DataFrame([
        {"CONFIG_KEY": "matching_version", "CONFIG_VALUE": MATCHING_VERSION},
        {"CONFIG_KEY": "reference_sheet", "CONFIG_VALUE": config.reference_sheet},
        {"CONFIG_KEY": "reference_header", "CONFIG_VALUE": config.reference_header},
        {"CONFIG_KEY": "reference_row_id_column", "CONFIG_VALUE": config.reference_row_id_column},
        {"CONFIG_KEY": "reference_company_column", "CONFIG_VALUE": config.reference_company_column},
        {"CONFIG_KEY": "reference_phone_columns", "CONFIG_VALUE": ", ".join(config.reference_phone_columns)},
        {"CONFIG_KEY": "reference_email_columns", "CONFIG_VALUE": ", ".join(config.reference_email_columns)},
        {"CONFIG_KEY": "source_sheet", "CONFIG_VALUE": config.source_sheet},
        {"CONFIG_KEY": "source_header", "CONFIG_VALUE": config.source_header},
        {"CONFIG_KEY": "source_phone_columns", "CONFIG_VALUE": ", ".join(config.source_phone_columns)},
        {"CONFIG_KEY": "source_email_columns", "CONFIG_VALUE": ", ".join(config.source_email_columns)},
        {"CONFIG_KEY": "source_company_column", "CONFIG_VALUE": config.source_company_column or ""},
    ])

    return {
        "summary": pd.DataFrame(summary_rows),
        "configuration": config_df,
        "reference_prepared": reference_prepared,
        "result": result,
        "matched": matched,
        "ambiguous": ambiguous,
        "not_found": not_found,
    }


def export_matching_excel(result: dict[str, pd.DataFrame], output_path: str | Path) -> Path:
    output = Path(output_path)
    if output.suffix.lower() != ".xlsx":
        output = output.with_suffix(".xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)

    sheets = {
        "Ringkasan Matching": result["summary"],
        "Konfigurasi": result["configuration"],
        "Hasil Matching": result["result"],
        "Data Cocok": result["matched"],
        "Data Ambigu": result["ambiguous"],
        "Data Tidak Ditemukan": result["not_found"],
        "Database Normalisasi": result["reference_prepared"],
    }
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, table in sheets.items():
            table.to_excel(writer, sheet_name=name[:31], index=False)
        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for cell in worksheet[1]:
                cell.font = Font(bold=True)
            for cells in worksheet.iter_cols():
                sample = [str(cell.value) if cell.value is not None else "" for cell in list(cells)[:250]]
                width = min(max(max((len(value) for value in sample), default=0) + 2, 10), 45)
                worksheet.column_dimensions[cells[0].column_letter].width = width
    return output


def safe_output_filename(source_name: str) -> str:
    stem = Path(sanitize_filename(source_name, "customer.xlsx")).stem
    return sanitize_filename(f"hasil_identifikasi_customer_{stem}.xlsx")

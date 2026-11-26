from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from openpyxl.styles import Font

from pipeline_etl_askrindo import (
    company_match_key,
    normalize_company_name,
    normalize_date_detailed,
    normalize_email_detailed,
    normalize_phone_detailed,
    normalize_person_name,
    person_name_match_key,
    sanitize_filename,
)

INTERACTION_PIPELINE_VERSION = "1.1.0"

DEFAULT_OUTPUT_SHEETS = [
    "Ringkasan Proses",
    "Konfigurasi",
    "Master Normalisasi",
    "Contact Centre Normalisasi",
    "Hasil Matching",
    "Ringkasan per Pelanggan",
    "Detail Riwayat Kontak",
    "Pelanggan Pernah Menghubungi",
    "Pelanggan Belum Menghubungi",
    "Data Ambigu",
    "Data Tidak Ditemukan",
    "Audit Matching",
]

MATCH_METHOD_LABELS = {
    "PHONE_EMAIL": "PHONE_EMAIL",
    "PHONE": "PHONE",
    "EMAIL": "EMAIL",
    "NAME_COMPANY": "NAME_COMPANY",
    "PHONE_COMPANY": "PHONE_COMPANY",
    "EMAIL_COMPANY": "EMAIL_COMPANY",
    "PHONE_NAME_COMPANY": "PHONE_NAME_COMPANY",
    "EMAIL_NAME_COMPANY": "EMAIL_NAME_COMPANY",
}

INTERACTION_METHODS = {
    "report_number": "Nomor pelaporan unik",
    "each_row": "Setiap baris Contact Centre",
    "date_channel_category": "Kombinasi tanggal + saluran + kategori",
}


@dataclass
class InteractionConfig:
    master_sheet: str | int
    master_header: int
    master_row_id_column: str
    master_customer_name_column: str
    master_company_name_column: str
    master_phone_columns: list[str] = field(default_factory=list)
    master_email_columns: list[str] = field(default_factory=list)
    contact_sheet: str | int = 0
    contact_header: int = 0
    contact_date_column: str | None = None
    contact_report_number_column: str | None = None
    contact_customer_name_column: str | None = None
    contact_company_name_column: str | None = None
    contact_phone_columns: list[str] = field(default_factory=list)
    contact_email_columns: list[str] = field(default_factory=list)
    contact_channel_column: str | None = None
    contact_category_column: str | None = None
    contact_service_type_column: str | None = None
    contact_status_column: str | None = None
    contact_description_column: str | None = None
    match_phone: bool = True
    match_email: bool = True
    match_name: bool = False
    match_company: bool = False
    allow_name_company_fallback: bool = False
    interaction_definition: str = "report_number"
    output_sheets: list[str] = field(default_factory=lambda: DEFAULT_OUTPUT_SHEETS.copy())
    master_filename: str = ""
    contact_filename: str = ""

    def __post_init__(self) -> None:
        self.output_sheets = list(dict.fromkeys(self.output_sheets))
        if self.interaction_definition not in INTERACTION_METHODS:
            raise ValueError("Definisi interaksi tidak dikenali.")


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


def _to_display_list(values: list[str]) -> str:
    return "; ".join(dict.fromkeys(v for v in values if v))


def _normalize_report_number(value: Any) -> Any:
    if pd.isna(value):
        return pd.NA
    text = str(value).strip()
    return pd.NA if text == "" else text


def _normalize_text(value: Any) -> str:
    return "" if pd.isna(value) else str(value).strip()


def validate_interaction_config(
    config: InteractionConfig,
    master_columns: list[str],
    contact_columns: list[str],
) -> list[str]:
    errors: list[str] = []
    master_set = set(master_columns)
    contact_set = set(contact_columns)

    if config.master_row_id_column not in master_set:
        errors.append(f"Kolom ROW_ID master tidak ditemukan: {config.master_row_id_column}")
    if config.master_customer_name_column not in master_set:
        errors.append(f"Kolom nama pelanggan master tidak ditemukan: {config.master_customer_name_column}")
    if config.master_company_name_column not in master_set:
        errors.append(f"Kolom nama perusahaan master tidak ditemukan: {config.master_company_name_column}")
    for column in config.master_phone_columns + config.master_email_columns:
        if column not in master_set:
            errors.append(f"Kolom master tidak ditemukan: {column}")
    if not config.master_phone_columns and not config.master_email_columns:
        errors.append("Master harus memiliki minimal satu kolom telepon atau email.")

    for column in config.contact_phone_columns + config.contact_email_columns:
        if column not in contact_set:
            errors.append(f"Kolom Contact Centre tidak ditemukan: {column}")
    if config.contact_date_column and config.contact_date_column not in contact_set:
        errors.append(f"Kolom tanggal Contact Centre tidak ditemukan: {config.contact_date_column}")
    if config.contact_report_number_column and config.contact_report_number_column not in contact_set:
        errors.append(f"Kolom nomor pelaporan Contact Centre tidak ditemukan: {config.contact_report_number_column}")
    if config.contact_customer_name_column and config.contact_customer_name_column not in contact_set:
        errors.append(f"Kolom nama pelanggan Contact Centre tidak ditemukan: {config.contact_customer_name_column}")
    if config.contact_company_name_column and config.contact_company_name_column not in contact_set:
        errors.append(f"Kolom nama perusahaan Contact Centre tidak ditemukan: {config.contact_company_name_column}")
    if not config.contact_phone_columns and not config.contact_email_columns:
        errors.append("Contact Centre harus memiliki minimal satu kolom telepon atau email.")
    if config.allow_name_company_fallback and not (config.match_name and config.match_company):
        errors.append("Fallback nama + perusahaan membutuhkan kedua opsi nama dan perusahaan diaktifkan.")
    return errors


def _prepare_master_database(
    master_df: pd.DataFrame,
    config: InteractionConfig,
) -> pd.DataFrame:
    work = master_df.copy().reset_index(drop=True)
    row_ids = work[config.master_row_id_column].astype("string").str.strip()
    if row_ids.isna().any() or row_ids.eq("").any():
        raise ValueError("ROW_ID master tidak boleh kosong.")
    duplicates = row_ids[row_ids.duplicated(keep=False)]
    if not duplicates.empty:
        sample = ", ".join(duplicates.drop_duplicates().head(10).tolist())
        raise ValueError(f"ROW_ID master harus unik. Duplikat ditemukan: {sample}")

    work["MASTER_ROW_ID"] = row_ids
    work["MASTER_CUSTOMER_NAME_ORIGINAL"] = work[config.master_customer_name_column]
    work["MASTER_CUSTOMER_NAME_NORMALIZED"] = work[config.master_customer_name_column].map(normalize_person_name)
    work["MASTER_CUSTOMER_NAME_MATCH_KEY"] = work[config.master_customer_name_column].map(person_name_match_key)
    work["MASTER_COMPANY_ORIGINAL"] = work[config.master_company_name_column]
    work["MASTER_COMPANY_NORMALIZED"] = work[config.master_company_name_column].map(normalize_company_name)
    work["MASTER_COMPANY_MATCH_KEY"] = work[config.master_company_name_column].map(company_match_key)

    master_phones: list[str] = []
    master_emails: list[str] = []
    for _, row in work.iterrows():
        phones: list[str] = []
        for column in config.master_phone_columns:
            phones.extend(_split_semicolon(normalize_phone_detailed(row.get(column)).normalized))
        emails: list[str] = []
        for column in config.master_email_columns:
            emails.extend(_split_semicolon(normalize_email_detailed(row.get(column)).normalized))
        master_phones.append(_to_display_list(phones))
        master_emails.append(_to_display_list(emails))
    work["MASTER_PHONE_VALUES"] = master_phones
    work["MASTER_EMAIL_VALUES"] = master_emails
    return work


def _build_master_indexes(master_prepared: pd.DataFrame) -> dict[str, Any]:
    phone_index: dict[str, set[str]] = {}
    email_index: dict[str, set[str]] = {}
    name_company_index: dict[str, set[str]] = {}
    row_lookup: dict[str, dict[str, Any]] = {}

    for _, row in master_prepared.iterrows():
        row_id = str(row["MASTER_ROW_ID"])
        row_lookup[row_id] = {
            "customer_name": row["MASTER_CUSTOMER_NAME_ORIGINAL"],
            "company_name": row["MASTER_COMPANY_ORIGINAL"],
            "company_key": row["MASTER_COMPANY_MATCH_KEY"],
            "name_key": row["MASTER_CUSTOMER_NAME_MATCH_KEY"],
        }
        for phone in _split_semicolon(row["MASTER_PHONE_VALUES"]):
            phone_index.setdefault(phone, set()).add(row_id)
        for email in _split_semicolon(row["MASTER_EMAIL_VALUES"]):
            email_index.setdefault(email, set()).add(row_id)
        combined_key = f"{row['MASTER_CUSTOMER_NAME_MATCH_KEY']}|{row['MASTER_COMPANY_MATCH_KEY']}"
        if row["MASTER_CUSTOMER_NAME_MATCH_KEY"] and row["MASTER_COMPANY_MATCH_KEY"]:
            name_company_index.setdefault(combined_key, set()).add(row_id)

    return {
        "phone": phone_index,
        "email": email_index,
        "name_company": name_company_index,
        "rows": row_lookup,
    }


def _prepare_contact_centre(
    contact_df: pd.DataFrame,
    config: InteractionConfig,
) -> pd.DataFrame:
    work = contact_df.copy().reset_index(drop=True)
    work.insert(0, "SOURCE_ROW_NUMBER", range(config.contact_header + 2, config.contact_header + 2 + len(work)))
    work.insert(0, "CONTACT_ROW_ID", [f"CONTACT-{i:07d}" for i in range(1, len(work) + 1)])
    
    work["SOURCE_CUSTOMER_NAME_ORIGINAL"] = work[config.contact_customer_name_column] if config.contact_customer_name_column else pd.NA
    work["SOURCE_CUSTOMER_NAME_NORMALIZED"] = work[config.contact_customer_name_column].map(normalize_person_name) if config.contact_customer_name_column else pd.NA
    work["SOURCE_CUSTOMER_NAME_MATCH_KEY"] = work[config.contact_customer_name_column].map(person_name_match_key) if config.contact_customer_name_column else ""
    work["SOURCE_COMPANY_ORIGINAL"] = work[config.contact_company_name_column] if config.contact_company_name_column else pd.NA
    work["SOURCE_COMPANY_NORMALIZED"] = work[config.contact_company_name_column].map(normalize_company_name) if config.contact_company_name_column else pd.NA
    work["SOURCE_COMPANY_MATCH_KEY"] = work[config.contact_company_name_column].map(company_match_key) if config.contact_company_name_column else ""

    source_phones: list[str] = []
    source_emails: list[str] = []
    for _, row in work.iterrows():
        phones: list[str] = []
        for column in config.contact_phone_columns:
            phones.extend(_split_semicolon(normalize_phone_detailed(row.get(column)).normalized))
        emails: list[str] = []
        for column in config.contact_email_columns:
            emails.extend(_split_semicolon(normalize_email_detailed(row.get(column)).normalized))
        source_phones.append(_to_display_list(phones))
        source_emails.append(_to_display_list(emails))
    work["SOURCE_PHONE_VALUES"] = source_phones
    work["SOURCE_EMAIL_VALUES"] = source_emails

    work["CONTACT_DATE_ORIGINAL"] = work[config.contact_date_column] if config.contact_date_column else pd.NA
    dates = []
    date_reasons = []
    for value in work["CONTACT_DATE_ORIGINAL"]:
        result = normalize_date_detailed(value, 1900, datetime.today().year)
        dates.append(result.normalized)
        date_reasons.append(result.reason)
    work["CONTACT_DATE_NORMALIZED"] = dates
    work["CONTACT_DATE_REASON"] = date_reasons

    work["REPORT_NUMBER"] = work[config.contact_report_number_column].map(_normalize_report_number) if config.contact_report_number_column else pd.NA
    work["CHANNEL"] = work[config.contact_channel_column] if config.contact_channel_column else pd.NA
    work["CATEGORY"] = work[config.contact_category_column] if config.contact_category_column else pd.NA
    work["SERVICE_TYPE"] = work[config.contact_service_type_column] if config.contact_service_type_column else pd.NA
    work["STATUS"] = work[config.contact_status_column] if config.contact_status_column else pd.NA
    work["DESCRIPTION"] = work[config.contact_description_column] if config.contact_description_column else pd.NA

    return work


def _candidates_from_values(values: list[str], index: dict[str, set[str]]) -> set[str]:
    candidates: set[str] = set()
    for value in values:
        candidates.update(index.get(value, set()))
    return candidates


def _narrow_candidates(
    candidates: set[str],
    indexes: dict[str, Any],
    name_key: str,
    company_key: str,
) -> tuple[set[str], bool, str, int]:
    if not candidates or not name_key or not company_key:
        return candidates, False, "", 0
    supported = {
        row_id for row_id in candidates
        if indexes["rows"][row_id]["name_key"] == name_key
        and indexes["rows"][row_id]["company_key"] == company_key
    }
    if len(supported) == 1:
        return supported, True, "NAME_COMPANY", 90
    return candidates, False, "", 0


def _match_one_record(
    source_phones: list[str],
    source_emails: list[str],
    source_name_key: str,
    source_company_key: str,
    indexes: dict[str, Any],
    config: InteractionConfig,
) -> dict[str, Any]:
    phone_candidates = (
        _candidates_from_values(source_phones, indexes["phone"])
        if config.match_phone else set()
    )
    email_candidates = (
        _candidates_from_values(source_emails, indexes["email"])
        if config.match_email else set()
    )
    phone_candidates.discard("")
    email_candidates.discard("")

    name_company_candidates: set[str] = set()
    if (
        config.match_name
        and config.match_company
        and source_name_key
        and source_company_key
    ):
        name_company_candidates = indexes["name_company"].get(
            f"{source_name_key}|{source_company_key}", set()
        )

    candidates: set[str] = set()
    method = "NONE"
    score = 0
    name_support = False
    company_support = False
    reason = ""

    if phone_candidates and email_candidates:
        intersection = phone_candidates & email_candidates
        if intersection:
            candidates = intersection
            method = "PHONE_EMAIL"
            score = 100
            reason = "Nomor telepon dan email mengarah ke kandidat master yang sama."
        else:
            # Kontak memberi dua identitas yang saling bertentangan. Jangan pilih salah satu.
            candidates = phone_candidates | email_candidates
            method = "PHONE_EMAIL_CONFLICT"
            score = 50
            reason = "Nomor telepon dan email mengarah ke ROW_ID master yang berbeda."
    elif phone_candidates:
        candidates = set(phone_candidates)
        method = "PHONE"
        score = 80
        reason = "Nomor telepon ditemukan pada master pelanggan."
    elif email_candidates:
        candidates = set(email_candidates)
        method = "EMAIL"
        score = 75
        reason = "Email ditemukan pada master pelanggan."

    # Nama + perusahaan hanya mempersempit kandidat, tidak pernah dipakai terpisah.
    if len(candidates) > 1 and name_company_candidates:
        supported = candidates & name_company_candidates
        if len(supported) == 1:
            candidates = supported
            name_support = True
            company_support = True
            if method == "PHONE":
                method = "PHONE_NAME_COMPANY"
                score = 90
            elif method == "EMAIL":
                method = "EMAIL_NAME_COMPANY"
                score = 85
            elif method in {"PHONE_EMAIL", "PHONE_EMAIL_CONFLICT"}:
                method = "PHONE_EMAIL_NAME_COMPANY"
                score = 95
            reason += " Nama pelanggan dan perusahaan mempersempit kandidat menjadi satu ROW_ID."

    if (
        not candidates
        and config.allow_name_company_fallback
        and config.match_name
        and config.match_company
        and name_company_candidates
    ):
        candidates = set(name_company_candidates)
        method = "NAME_COMPANY"
        score = 60
        name_support = True
        company_support = True
        reason = "Telepon/email tidak menemukan kandidat; exact match nama dan perusahaan digunakan sebagai fallback."

    candidate_ids = " | ".join(sorted(candidates))
    phone_ids = " | ".join(sorted(phone_candidates))
    email_ids = " | ".join(sorted(email_candidates))
    name_company_ids = " | ".join(sorted(name_company_candidates))

    common = {
        "PHONE_CANDIDATES": phone_ids,
        "EMAIL_CANDIDATES": email_ids,
        "NAME_COMPANY_CANDIDATES": name_company_ids,
        "FINAL_CANDIDATES": candidate_ids,
    }

    if len(candidates) == 1:
        row_id = next(iter(candidates))
        info = indexes["rows"].get(row_id, {})
        return {
            **common,
            "MATCHED_CUSTOMER_ROW_ID": row_id,
            "MATCHED_CUSTOMER_NAME": info.get("customer_name", pd.NA),
            "MATCHED_COMPANY_NAME": info.get("company_name", pd.NA),
            "MATCH_STATUS": "MATCHED",
            "MATCH_METHOD": method,
            "MATCH_SCORE": score,
            "MATCH_CANDIDATE_COUNT": 1,
            "MATCH_CANDIDATE_ROW_IDS": row_id,
            "MATCH_REASON": reason or "Kontak cocok ke satu ROW_ID master pelanggan.",
            "NAME_USED_AS_SUPPORT": name_support,
            "COMPANY_USED_AS_SUPPORT": company_support,
        }

    if len(candidates) > 1:
        return {
            **common,
            "MATCHED_CUSTOMER_ROW_ID": pd.NA,
            "MATCHED_CUSTOMER_NAME": pd.NA,
            "MATCHED_COMPANY_NAME": pd.NA,
            "MATCH_STATUS": "AMBIGUOUS",
            "MATCH_METHOD": method,
            "MATCH_SCORE": 50,
            "MATCH_CANDIDATE_COUNT": len(candidates),
            "MATCH_CANDIDATE_ROW_IDS": candidate_ids,
            "MATCH_REASON": reason or "Kontak cocok ke beberapa ROW_ID dan tidak dapat dipersempit secara unik.",
            "NAME_USED_AS_SUPPORT": name_support,
            "COMPANY_USED_AS_SUPPORT": company_support,
        }

    return {
        **common,
        "MATCHED_CUSTOMER_ROW_ID": pd.NA,
        "MATCHED_CUSTOMER_NAME": pd.NA,
        "MATCHED_COMPANY_NAME": pd.NA,
        "MATCH_STATUS": "NOT_FOUND",
        "MATCH_METHOD": "NONE",
        "MATCH_SCORE": 0,
        "MATCH_CANDIDATE_COUNT": 0,
        "MATCH_CANDIDATE_ROW_IDS": "",
        "MATCH_REASON": "Telepon, email, serta kombinasi nama dan perusahaan tidak menemukan ROW_ID master.",
        "NAME_USED_AS_SUPPORT": False,
        "COMPANY_USED_AS_SUPPORT": False,
    }

def _interaction_key_for_row(
    row: pd.Series,
    config: InteractionConfig,
) -> str:
    matched_id = str(row["MATCHED_CUSTOMER_ROW_ID"])
    if row["MATCH_STATUS"] != "MATCHED":
        return f"UNMATCHED|{row['CONTACT_ROW_ID']}"

    if config.interaction_definition == "each_row":
        return f"{matched_id}|ROW|{row['CONTACT_ROW_ID']}"
    if config.interaction_definition == "report_number":
        report = _normalize_text(row["REPORT_NUMBER"])
        if report:
            return f"{matched_id}|REPORT|{report}"
        return f"{matched_id}|ROW|{row['CONTACT_ROW_ID']}"
    if config.interaction_definition == "date_channel_category":
        date_value = _normalize_text(row["CONTACT_DATE_NORMALIZED"]) or _normalize_text(row["CONTACT_DATE_ORIGINAL"])
        channel = _normalize_text(row["CHANNEL"])
        category = _normalize_text(row["CATEGORY"])
        return f"{matched_id}|DATE|{date_value}|CHANNEL|{channel}|CATEGORY|{category}"
    return f"{matched_id}|ROW|{row['CONTACT_ROW_ID']}"


def _format_date_display(value: Any, normalized: Any) -> str:
    if not pd.isna(normalized) and normalized:
        return str(normalized)
    if pd.isna(value):
        return ""
    return str(value).strip()


def _compute_customer_summary(
    matched: pd.DataFrame,
    master_prepared: pd.DataFrame,
) -> pd.DataFrame:
    if matched.empty:
        summary = master_prepared.copy()
        summary = summary.assign(
            TOTAL_CONTACTS=0,
            UNIQUE_REPORT_COUNT=0,
            FIRST_CONTACT_DATE=pd.NA,
            LAST_CONTACT_DATE=pd.NA,
            DAYS_SINCE_LAST_CONTACT=pd.NA,
            CONTACT_DATES="",
            CONTACT_CHANNELS="",
            SERVICE_TYPES="",
            CATEGORIES="",
            REPORT_NUMBERS="",
            MATCHED_CONTACT_ROWS="",
            EVER_CONTACTED=False,
        )
        return summary[
            [
                "MASTER_ROW_ID",
                "MASTER_CUSTOMER_NAME_ORIGINAL",
                "MASTER_COMPANY_ORIGINAL",
                "MASTER_PHONE_VALUES",
                "MASTER_EMAIL_VALUES",
                "TOTAL_CONTACTS",
                "UNIQUE_REPORT_COUNT",
                "FIRST_CONTACT_DATE",
                "LAST_CONTACT_DATE",
                "DAYS_SINCE_LAST_CONTACT",
                "CONTACT_DATES",
                "CONTACT_CHANNELS",
                "SERVICE_TYPES",
                "CATEGORIES",
                "REPORT_NUMBERS",
                "MATCHED_CONTACT_ROWS",
                "EVER_CONTACTED",
            ]
        ]

    agg = []
    for row_id, group in matched.groupby("MATCHED_CUSTOMER_ROW_ID", sort=False):
        dates = [d for d in group["CONTACT_DATE_NORMALIZED"].tolist() if not pd.isna(d) and d]
        bool_dates = [pd.to_datetime(d) for d in dates if d]
        first_date = min(bool_dates) if bool_dates else None
        last_date = max(bool_dates) if bool_dates else None
        if last_date is not None:
            days_since_last = int((pd.Timestamp.now().normalize() - last_date.normalize()).days)
        else:
            days_since_last = pd.NA
        report_numbers = [str(item).strip() for item in group["REPORT_NUMBER"].tolist() if not pd.isna(item) and str(item).strip()]
        interaction_keys = group["INTERACTION_KEY"].dropna().unique().tolist()
        channels = [str(item).strip() for item in group["CHANNEL"].tolist() if not pd.isna(item) and str(item).strip()]
        services = [str(item).strip() for item in group["SERVICE_TYPE"].tolist() if not pd.isna(item) and str(item).strip()]
        categories = [str(item).strip() for item in group["CATEGORY"].tolist() if not pd.isna(item) and str(item).strip()]
        contact_rows = [str(item).strip() for item in group["CONTACT_ROW_ID"].tolist() if not pd.isna(item)]
        date_display = [
            _format_date_display(orig, norm)
            for orig, norm in zip(group["CONTACT_DATE_ORIGINAL"].tolist(), group["CONTACT_DATE_NORMALIZED"].tolist())
        ]
        agg.append({
            "MATCHED_CUSTOMER_ROW_ID": row_id,
            "TOTAL_CONTACTS": len(set(interaction_keys)),
            "UNIQUE_REPORT_COUNT": len(set(report_numbers)),
            "FIRST_CONTACT_DATE": first_date.strftime("%m/%d/%Y") if first_date is not None else pd.NA,
            "LAST_CONTACT_DATE": last_date.strftime("%m/%d/%Y") if last_date is not None else pd.NA,
            "DAYS_SINCE_LAST_CONTACT": days_since_last,
            "CONTACT_DATES": _to_display_list(date_display),
            "CONTACT_CHANNELS": _to_display_list(channels),
            "SERVICE_TYPES": _to_display_list(services),
            "CATEGORIES": _to_display_list(categories),
            "REPORT_NUMBERS": _to_display_list(report_numbers),
            "MATCHED_CONTACT_ROWS": _to_display_list(contact_rows),
            "EVER_CONTACTED": True,
        })

    summary = pd.DataFrame(agg)
    master_summary = master_prepared[[
        "MASTER_ROW_ID",
        "MASTER_CUSTOMER_NAME_ORIGINAL",
        "MASTER_COMPANY_ORIGINAL",
        "MASTER_PHONE_VALUES",
        "MASTER_EMAIL_VALUES",
    ]].copy()
    result = master_summary.merge(summary, how="left", left_on="MASTER_ROW_ID", right_on="MATCHED_CUSTOMER_ROW_ID")
    result["TOTAL_CONTACTS"] = result["TOTAL_CONTACTS"].fillna(0).astype(int)
    result["UNIQUE_REPORT_COUNT"] = result["UNIQUE_REPORT_COUNT"].fillna(0).astype(int)
    result["FIRST_CONTACT_DATE"] = result["FIRST_CONTACT_DATE"].where(result["TOTAL_CONTACTS"] > 0, pd.NA)
    result["LAST_CONTACT_DATE"] = result["LAST_CONTACT_DATE"].where(result["TOTAL_CONTACTS"] > 0, pd.NA)
    result["DAYS_SINCE_LAST_CONTACT"] = result["DAYS_SINCE_LAST_CONTACT"].where(result["TOTAL_CONTACTS"] > 0, pd.NA)
    result["CONTACT_DATES"] = result["CONTACT_DATES"].fillna("")
    result["CONTACT_CHANNELS"] = result["CONTACT_CHANNELS"].fillna("")
    result["SERVICE_TYPES"] = result["SERVICE_TYPES"].fillna("")
    result["CATEGORIES"] = result["CATEGORIES"].fillna("")
    result["REPORT_NUMBERS"] = result["REPORT_NUMBERS"].fillna("")
    result["MATCHED_CONTACT_ROWS"] = result["MATCHED_CONTACT_ROWS"].fillna("")
    result["EVER_CONTACTED"] = result["EVER_CONTACTED"].astype("boolean").fillna(False).astype(bool)
    result = result[[
        "MASTER_ROW_ID",
        "MASTER_CUSTOMER_NAME_ORIGINAL",
        "MASTER_COMPANY_ORIGINAL",
        "MASTER_PHONE_VALUES",
        "MASTER_EMAIL_VALUES",
        "TOTAL_CONTACTS",
        "UNIQUE_REPORT_COUNT",
        "FIRST_CONTACT_DATE",
        "LAST_CONTACT_DATE",
        "DAYS_SINCE_LAST_CONTACT",
        "CONTACT_DATES",
        "CONTACT_CHANNELS",
        "SERVICE_TYPES",
        "CATEGORIES",
        "REPORT_NUMBERS",
        "MATCHED_CONTACT_ROWS",
        "EVER_CONTACTED",
    ]]
    return result


def run_interaction_pipeline(
    master_df: pd.DataFrame,
    contact_df: pd.DataFrame,
    config: InteractionConfig,
) -> dict[str, pd.DataFrame]:
    errors = validate_interaction_config(config, list(master_df.columns), list(contact_df.columns))
    if errors:
        raise ValueError(" | ".join(errors))

    master_prepared = _prepare_master_database(master_df, config)
    indexes = _build_master_indexes(master_prepared)
    contact_prepared = _prepare_contact_centre(contact_df, config)

    match_rows: list[dict[str, Any]] = []
    for _, row in contact_prepared.iterrows():
        phones = [item for item in _split_semicolon(row["SOURCE_PHONE_VALUES"]) if item]
        emails = [item for item in _split_semicolon(row["SOURCE_EMAIL_VALUES"]) if item]
        info = _match_one_record(
            phones,
            emails,
            row["SOURCE_CUSTOMER_NAME_MATCH_KEY"],
            row["SOURCE_COMPANY_MATCH_KEY"],
            indexes,
            config,
        )
        match_rows.append(info)

    matched = pd.concat([contact_prepared, pd.DataFrame(match_rows)], axis=1)
    matched["MATCHED_PHONE"] = matched["SOURCE_PHONE_VALUES"]
    matched["MATCHED_EMAIL"] = matched["SOURCE_EMAIL_VALUES"]
    matched["MATCH_CANDIDATE_COUNT"] = matched["MATCH_CANDIDATE_COUNT"].fillna(0).astype(int)
    matched["MATCH_CANDIDATE_ROW_IDS"] = matched["MATCH_CANDIDATE_ROW_IDS"].fillna("")
    matched["MATCHED_CUSTOMER_ROW_ID"] = matched["MATCHED_CUSTOMER_ROW_ID"].astype("string")

    matched["INTERACTION_KEY"] = matched.apply(lambda row: _interaction_key_for_row(row, config), axis=1)

    # Detail riwayat hanya memuat kontak yang sudah terhubung ke satu ROW_ID master.
    detail = matched[matched["MATCH_STATUS"] == "MATCHED"].copy()
    detail["__CONTACT_DATE_SORT__"] = pd.to_datetime(
        detail["CONTACT_DATE_NORMALIZED"], errors="coerce"
    )
    detail = detail.sort_values(
        by=["MATCHED_CUSTOMER_ROW_ID", "__CONTACT_DATE_SORT__", "SOURCE_ROW_NUMBER"],
        na_position="last",
        kind="stable",
    )
    detail["CONTACT_SEQUENCE"] = detail.groupby("MATCHED_CUSTOMER_ROW_ID", dropna=False).cumcount() + 1
    detail["PREVIOUS_CONTACT_DATE"] = pd.NA
    detail["DAYS_SINCE_PREVIOUS_CONTACT"] = pd.NA

    for customer_id, group in detail.groupby("MATCHED_CUSTOMER_ROW_ID", sort=False, dropna=False):
        previous_date = None
        for idx, row in group.iterrows():
            if not pd.isna(row["CONTACT_DATE_NORMALIZED"]) and row["CONTACT_DATE_NORMALIZED"]:
                current_date = pd.to_datetime(row["CONTACT_DATE_NORMALIZED"], errors="coerce")
            else:
                current_date = None
            if previous_date is None or current_date is None:
                detail.at[idx, "PREVIOUS_CONTACT_DATE"] = pd.NA
                detail.at[idx, "DAYS_SINCE_PREVIOUS_CONTACT"] = pd.NA
            else:
                detail.at[idx, "PREVIOUS_CONTACT_DATE"] = previous_date.strftime("%m/%d/%Y")
                detail.at[idx, "DAYS_SINCE_PREVIOUS_CONTACT"] = int((current_date - previous_date).days)
            if current_date is not None:
                previous_date = current_date

    detail = detail.drop(columns=["__CONTACT_DATE_SORT__"], errors="ignore")
    customer_summary = _compute_customer_summary(detail, master_prepared)

    matched_rows = matched[matched["MATCH_STATUS"] == "MATCHED"].copy()
    ambiguous_rows = matched[matched["MATCH_STATUS"] == "AMBIGUOUS"].copy()
    not_found_rows = matched[matched["MATCH_STATUS"] == "NOT_FOUND"].copy()

    customer_ever = customer_summary[customer_summary["EVER_CONTACTED"] == True].copy()
    customer_never = customer_summary[customer_summary["EVER_CONTACTED"] == False].copy()

    total_interactions = detail["INTERACTION_KEY"].nunique()
    matched_count = len(matched_rows)
    ambiguous_count = len(ambiguous_rows)
    not_found_count = len(not_found_rows)
    total_master = len(master_prepared)
    ever_contacted_count = len(customer_ever)
    never_contacted_count = len(customer_never)
    completeness_matching = round(matched_count / max(len(detail), 1) * 100, 2)

    first_date = pd.to_datetime(matched_rows["CONTACT_DATE_NORMALIZED"], errors="coerce")
    first_date = first_date.min() if not first_date.empty else pd.NaT
    last_date = pd.to_datetime(matched_rows["CONTACT_DATE_NORMALIZED"], errors="coerce")
    last_date = last_date.max() if not last_date.empty else pd.NaT
    most_contacts = customer_summary.sort_values(by="TOTAL_CONTACTS", ascending=False).head(1)
    most_contacts_label = ""
    if not most_contacts.empty and int(most_contacts.iloc[0]["TOTAL_CONTACTS"]) > 0:
        most_contacts_label = f"{most_contacts.iloc[0]['MASTER_ROW_ID']} ({most_contacts.iloc[0]['TOTAL_CONTACTS']})"

    summary_rows = [
        {"METRIC": "Nama file master", "VALUE": config.master_filename},
        {"METRIC": "Nama file Contact Centre", "VALUE": config.contact_filename},
        {"METRIC": "Sheet master", "VALUE": config.master_sheet},
        {"METRIC": "Sheet Contact Centre", "VALUE": config.contact_sheet},
        {"METRIC": "Header master", "VALUE": config.master_header},
        {"METRIC": "Header Contact Centre", "VALUE": config.contact_header},
        {"METRIC": "Waktu proses", "VALUE": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
        {"METRIC": "Versi pipeline", "VALUE": INTERACTION_PIPELINE_VERSION},
        {"METRIC": "Jumlah master pelanggan", "VALUE": total_master},
        {"METRIC": "Jumlah baris Contact Centre", "VALUE": len(detail)},
        {"METRIC": "Jumlah MATCHED", "VALUE": matched_count},
        {"METRIC": "Jumlah AMBIGUOUS", "VALUE": ambiguous_count},
        {"METRIC": "Jumlah NOT_FOUND", "VALUE": not_found_count},
        {"METRIC": "Jumlah pelanggan pernah menghubungi", "VALUE": ever_contacted_count},
        {"METRIC": "Jumlah pelanggan belum menghubungi", "VALUE": never_contacted_count},
        {"METRIC": "Total interaksi", "VALUE": total_interactions},
        {"METRIC": "Kompleteness matching (%)", "VALUE": completeness_matching},
        {"METRIC": "Rata-rata interaksi per pelanggan", "VALUE": round(customer_summary["TOTAL_CONTACTS"].mean() if not customer_summary.empty else 0, 2)},
        {"METRIC": "Median interaksi per pelanggan", "VALUE": float(customer_summary["TOTAL_CONTACTS"].median() if not customer_summary.empty else 0)},
        {"METRIC": "Pelanggan dengan interaksi terbanyak", "VALUE": most_contacts_label},
        {"METRIC": "Tanggal kontak paling awal", "VALUE": first_date.strftime("%m/%d/%Y") if pd.notna(first_date) else ""},
        {"METRIC": "Tanggal kontak paling akhir", "VALUE": last_date.strftime("%m/%d/%Y") if pd.notna(last_date) else ""},
        {"METRIC": "Status proses", "VALUE": "Sukses"},
    ]

    configuration = pd.DataFrame([
        {"CONFIG_KEY": "master_sheet", "CONFIG_VALUE": config.master_sheet},
        {"CONFIG_KEY": "master_header", "CONFIG_VALUE": config.master_header},
        {"CONFIG_KEY": "master_row_id_column", "CONFIG_VALUE": config.master_row_id_column},
        {"CONFIG_KEY": "master_customer_name_column", "CONFIG_VALUE": config.master_customer_name_column},
        {"CONFIG_KEY": "master_company_name_column", "CONFIG_VALUE": config.master_company_name_column},
        {"CONFIG_KEY": "master_phone_columns", "CONFIG_VALUE": "; ".join(config.master_phone_columns)},
        {"CONFIG_KEY": "master_email_columns", "CONFIG_VALUE": "; ".join(config.master_email_columns)},
        {"CONFIG_KEY": "contact_sheet", "CONFIG_VALUE": config.contact_sheet},
        {"CONFIG_KEY": "contact_header", "CONFIG_VALUE": config.contact_header},
        {"CONFIG_KEY": "contact_date_column", "CONFIG_VALUE": config.contact_date_column or ""},
        {"CONFIG_KEY": "contact_report_number_column", "CONFIG_VALUE": config.contact_report_number_column or ""},
        {"CONFIG_KEY": "contact_customer_name_column", "CONFIG_VALUE": config.contact_customer_name_column or ""},
        {"CONFIG_KEY": "contact_company_name_column", "CONFIG_VALUE": config.contact_company_name_column or ""},
        {"CONFIG_KEY": "contact_phone_columns", "CONFIG_VALUE": "; ".join(config.contact_phone_columns)},
        {"CONFIG_KEY": "contact_email_columns", "CONFIG_VALUE": "; ".join(config.contact_email_columns)},
        {"CONFIG_KEY": "contact_channel_column", "CONFIG_VALUE": config.contact_channel_column or ""},
        {"CONFIG_KEY": "contact_category_column", "CONFIG_VALUE": config.contact_category_column or ""},
        {"CONFIG_KEY": "contact_service_type_column", "CONFIG_VALUE": config.contact_service_type_column or ""},
        {"CONFIG_KEY": "contact_status_column", "CONFIG_VALUE": config.contact_status_column or ""},
        {"CONFIG_KEY": "contact_description_column", "CONFIG_VALUE": config.contact_description_column or ""},
        {"CONFIG_KEY": "match_phone", "CONFIG_VALUE": config.match_phone},
        {"CONFIG_KEY": "match_email", "CONFIG_VALUE": config.match_email},
        {"CONFIG_KEY": "match_name", "CONFIG_VALUE": config.match_name},
        {"CONFIG_KEY": "match_company", "CONFIG_VALUE": config.match_company},
        {"CONFIG_KEY": "allow_name_company_fallback", "CONFIG_VALUE": config.allow_name_company_fallback},
        {"CONFIG_KEY": "interaction_definition", "CONFIG_VALUE": INTERACTION_METHODS[config.interaction_definition]},
        {"CONFIG_KEY": "output_sheets", "CONFIG_VALUE": "; ".join(config.output_sheets)},
    ])

    return {
        "summary": pd.DataFrame(summary_rows),
        "configuration": configuration,
        "master_normalized": master_prepared,
        "contact_normalized": contact_prepared,
        "result": matched,
        "detail": detail,
        "matched": matched_rows,
        "ambiguous": ambiguous_rows,
        "not_found": not_found_rows,
        "customer_summary": customer_summary,
        "customer_ever": customer_ever,
        "customer_never": customer_never,
        "audit": _build_audit_table(contact_prepared, match_rows),
        "output_sheets": config.output_sheets,
    }


def _build_audit_table(contact_prepared: pd.DataFrame, match_rows: list[dict[str, Any]]) -> pd.DataFrame:
    audit = contact_prepared[[
        "CONTACT_ROW_ID",
        "SOURCE_PHONE_VALUES",
        "SOURCE_EMAIL_VALUES",
        "SOURCE_CUSTOMER_NAME_ORIGINAL",
        "SOURCE_COMPANY_ORIGINAL",
    ]].copy()
    match_df = pd.DataFrame(match_rows)
    audit = pd.concat([audit.reset_index(drop=True), match_df.reset_index(drop=True)], axis=1)
    required = [
        "CONTACT_ROW_ID",
        "SOURCE_PHONE_VALUES",
        "SOURCE_EMAIL_VALUES",
        "SOURCE_CUSTOMER_NAME_ORIGINAL",
        "SOURCE_COMPANY_ORIGINAL",
        "PHONE_CANDIDATES",
        "EMAIL_CANDIDATES",
        "NAME_COMPANY_CANDIDATES",
        "FINAL_CANDIDATES",
        "MATCHED_CUSTOMER_ROW_ID",
        "MATCH_STATUS",
        "MATCH_METHOD",
        "MATCH_SCORE",
        "MATCH_REASON",
    ]
    for column in required:
        if column not in audit.columns:
            audit[column] = ""
    return audit[required]

def _safe_sheet_name(name: str) -> str:
    safe = re.sub(r"[\\/*?:\[\]]", "_", name)
    return safe[:31]


def export_interaction_excel(result: dict[str, pd.DataFrame], output_path: str | Path) -> Path:
    output = Path(output_path)
    if output.suffix.lower() != ".xlsx":
        output = output.with_suffix(".xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)

    sheets = {
        "Ringkasan Proses": result["summary"],
        "Konfigurasi": result["configuration"],
        "Master Normalisasi": result["master_normalized"],
        "Contact Centre Normalisasi": result["contact_normalized"],
        "Hasil Matching": result["result"],
        "Ringkasan per Pelanggan": result["customer_summary"],
        "Detail Riwayat Kontak": result["detail"],
        "Pelanggan Pernah Menghubungi": result["customer_ever"],
        "Pelanggan Belum Menghubungi": result["customer_never"],
        "Data Ambigu": result["ambiguous"],
        "Data Tidak Ditemukan": result["not_found"],
        "Audit Matching": result["audit"],
    }
    selected_sheets = result.get("output_sheets", list(sheets.keys()))
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name in selected_sheets:
            table = sheets.get(name)
            if table is None:
                continue
            table.to_excel(writer, sheet_name=_safe_sheet_name(name), index=False)
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
    stem = Path(sanitize_filename(source_name, "contact_centre.xlsx")).stem
    return sanitize_filename(f"riwayat_interaksi_pelanggan_{stem}.xlsx")

from __future__ import annotations

import argparse
import io
import json
import re
import time
import unicodedata
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Iterable, Literal

import numpy as np
import pandas as pd
from openpyxl.styles import Font

PIPELINE_VERSION = "6.2.0"

PipelineMode = Literal["Generic"]

REPORT_REGISTRY: dict[str, str] = {
    "summary": "Ringkasan Proses",
    "config": "Konfigurasi",
    "raw": "Data Asli",
    "normalized": "Data Normalisasi",
    "clean": "Data Bersih",
    "removed": "Data Dihapus",
    "audit": "Audit Perubahan",
    "completeness": "Completeness Kolom",
    "missing": "Data Kosong",
    "validity": "Validity Kolom",
    "invalid": "Data Tidak Valid",
    "duplicates": "Data Duplikat",
    "duplicate_report_number": "Duplikat No Pelaporan",
    "report_name_conflict": "Konflik Nama Pelaporan",
    "legacy_duplicates": "Duplikat Responden Lama",
    "multi_policy": "Multi Policy",
    "multi_policy_detail": "Multi Policy Detail",
    "duplicate_policy_summary": "Ringkasan Duplikat Policy",
    "relations": "Relasi Satu ke Banyak",
    "header_mapping": "Pemetaan Header",
    "removal_log": "Removal Log",
}

DEFAULT_OUTPUT_SHEETS: tuple[str, ...] = (
    "summary", "config", "normalized", "clean", "completeness", "validity"
)

ColumnType = Literal[
    "text",
    "person_name",
    "company_name",
    "phone",
    "email",
    "date",
    "month",
    "year",
    "identity",
    "nik",
    "npwp",
    "policy_number",
    "report_number",
    "category",
    "numeric",
    "free",
]

COLUMN_TYPE_LABELS: dict[str, str] = {
    "text": "Teks umum",
    "person_name": "Nama orang",
    "company_name": "Nama perusahaan",
    "phone": "Nomor telepon",
    "email": "Email",
    "date": "Tanggal",
    "month": "Bulan",
    "year": "Tahun",
    "identity": "Nomor identitas",
    "nik": "NIK",
    "npwp": "NPWP",
    "policy_number": "Nomor polis",
    "report_number": "Nomor pelaporan",
    "category": "Kategori",
    "numeric": "Kolom numerik",
    "free": "Bebas tanpa normalisasi",
}

DEFAULT_PLACEHOLDERS: tuple[str, ...] = (
    "NULL",
    "NONE",
    "N/A",
    "NA",
    "NAN",
    "KOSONG",
    "MIGRATION",
    "MIGRASI",
    "UNKNOWN",
    "TIDAK ADA",
    "BELUM ADA",
    "TIDAK DIKETAHUI",
    "NOT AVAILABLE",
    "NO DATA",
    "EMPTY",
    "<M-ADDRESS KOSONG>",
    "M-ADDRESS KOSONG",
    "M ADDRESS KOSONG",
    "MIGRATION ADDRESS KOSONG",
    "MIGRATION ADDRES KOSONG",
    "MIGRASI ALAMAT KOSONG",
    "ALAMAT KOSONG",
    "ADDRESS KOSONG",
    "-",
    "--",
    ".",
    "#N/A",
    "#VALUE!",
    "#REF!",
    "#DIV/0!",
    "#NAME?",
    "#NUM!",
    "#NULL!",
)

# Alias lama dipertahankan sebagai basis rekomendasi mapping. Alias tidak pernah
# diterapkan diam-diam; pengguna harus meninjau mapping pada antarmuka.
HEADER_ALIASES: dict[str, list[str]] = {
    "Nomor Pelaporan": [
        "Nomor Pelaporan", "No Pelaporan", "No. Pelaporan", "Nomor Laporan",
        "No Laporan", "Reporting Number", "Ticket ID", "Nomor Tiket",
    ],
    "Tanggal": [
        "Tanggal", "Tanggal Pelaporan", "Tgl Pelaporan", "Tanggal Laporan",
        "Date", "Reporting Date", "Created Date",
    ],
    "Bulan": ["Bulan", "Month", "Bulan Pelaporan"],
    "Tahun": ["Tahun", "Year", "Tahun Pelaporan"],
    "CUSTOMER_ID": [
        "CUSTOMER_ID", "Customer ID", "ID Customer", "ID Pelanggan",
        "CustomerID", "Kode Pelanggan",
    ],
    "Nama Pelanggan": ["Nama Pelanggan", "Customer Name", "Nama Customer", "Pelanggan"],
    "Nama Kontak": ["Nama Kontak", "Contact Name", "Nama PIC", "PIC", "Contact Person"],
    "Nama Perusahaan": ["Nama Perusahaan", "Company Name", "Perusahaan", "Nama Instansi"],
    "Kanwil": ["Kanwil", "Kantor Wilayah", "Regional Office", "Wilayah"],
    "Kantor Cabang": ["Kantor Cabang", "Cabang", "Branch", "Branch Office", "KC"],
    "Nomor Hp Pelanggan": [
        "Nomor Hp Pelanggan", "Nomor HP Pelanggan", "No HP Pelanggan",
        "No. HP Pelanggan", "Handphone", "No Handphone", "Nomor Handphone",
        "Nomor Telepon", "No Telepon", "Telepon", "Phone", "Mobile",
    ],
    "WhatsApp": [
        "WhatsApp", "Whatsapp", "WA", "Nomor WhatsApp", "No WhatsApp",
        "Nomor WA", "No WA",
    ],
    "Email pelanggan": [
        "Email pelanggan", "Email Pelanggan", "Email", "E-mail",
        "Alamat Email", "Customer Email",
    ],
    "Nama Pemegang Polis": ["Nama Pemegang Polis", "Policy Holder", "Pemegang Polis"],
    "Nomor Polis": ["Nomor Polis", "No Polis", "No. Polis", "Policy Number", "POLICY_NO"],
    "POLICY_ID": ["POLICY_ID", "Policy ID", "ID Policy", "ID Polis"],
    "Produk Yang Digunakan": [
        "Produk Yang Digunakan", "Produk Digunakan", "Produk yang digunakan",
        "Used Product", "Produk Pelanggan",
    ],
    "Jenis Layanan": ["Jenis Layanan", "Service Type", "Layanan"],
    "Kategori": ["Kategori", "Category"],
    "COB": ["COB", "Class of Business"],
    "Produk": ["Produk", "Product"],
    "Keterangan": ["Keterangan", "Description", "Catatan", "Remarks"],
    "Tindak Lanjut": ["Tindak Lanjut", "Follow Up", "Follow-up"],
    "Status": ["Status", "Status Pelaporan"],
    "Saluran": ["Saluran", "Channel", "Media"],
    "Sifat Pelaporan": ["Sifat Pelaporan", "Jenis Pelaporan", "Reporting Nature"],
    "Agent": ["Agent", "Agen", "Petugas"],
    "KTP_DEBITUR": ["KTP_DEBITUR", "KTP Debitur", "NIK Debitur", "NIK", "Nomor KTP", "No KTP", "IDNO"],
    "NPWP_DEBITUR": ["NPWP_DEBITUR", "NPWP Debitur", "NPWP", "Nomor NPWP", "No NPWP"],
    "ALAMAT_DEBITUR": ["ALAMAT_DEBITUR", "Alamat Debitur", "Alamat"],
    "KELURAHAN_DEBITUR": ["KELURAHAN_DEBITUR", "Kelurahan Debitur", "Kelurahan"],
    "KECAMATAN_DEBITUR": ["KECAMATAN_DEBITUR", "Kecamatan Debitur", "Kecamatan"],
    "KOTA_DEBITUR": ["KOTA_DEBITUR", "Kota Debitur"],
    "PROVINSI_DEBITUR": ["PROVINSI_DEBITUR", "Provinsi Debitur"],
    "KOTA": ["KOTA", "Kota"],
    "PROVINSI": ["PROVINSI", "Provinsi"],
    "DEBITUR_ID": ["DEBITUR_ID", "Debitur ID", "ID Debitur"],
    "BRANCH_ID": ["BRANCH_ID", "Branch ID", "ID Cabang"],
    "POSTING_DATE": ["POSTING_DATE", "Posting Date", "Tanggal Posting"],
    "BANK_INDUK": ["BANK_INDUK", "Bank Induk"],
    "ID_BANK_CABANG": ["ID_BANK_CABANG", "ID Bank Cabang"],
    "BANK_CABANG": ["BANK_CABANG", "Bank Cabang", "Nama Bank Cabang"],
    "ALAMAT_BANK_CABANG": ["ALAMAT_BANK_CABANG", "Alamat Bank Cabang"],
    "KODE_POS": ["KODE_POS", "Kode Pos"],
    "KATEGORI_BANK": ["KATEGORI_BANK", "Kategori Bank"],
    "TIPE_DEBITUR": ["TIPE_DEBITUR", "Tipe Debitur", "Jenis Debitur"],
    "NAMA_DEBITUR": ["NAMA_DEBITUR", "Nama Debitur"],
    "JENIS_KELAMIN_DEBITUR": ["JENIS_KELAMIN_DEBITUR", "Jenis Kelamin Debitur", "Gender Debitur"],
    "TEMPAT_LAHIR_DEBITUR": ["TEMPAT_LAHIR_DEBITUR", "Tempat Lahir Debitur"],
    "TANGGAL_LAHIR_DEBITUR": ["TANGGAL_LAHIR_DEBITUR", "Tanggal Lahir Debitur"],
    "KODE_POS_DEBITUR": ["KODE_POS_DEBITUR", "Kode Pos Debitur"],
    "NEGARA_DEBITUR": ["NEGARA_DEBITUR", "Negara Debitur"],
    "REPORTING_DATE": ["REPORTING_DATE", "Reporting Date", "Tanggal Reporting"],
}


def _header_key(value: Any) -> str:
    """Membuat kunci pembanding header yang tahan spasi, kapital, dan simbol."""
    text = "" if value is None else str(value)
    text = re.sub(r"\.\d+$", "", text.strip())
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.strip().casefold()
    return re.sub(r"[^a-z0-9]+", "", text)


ALIAS_LOOKUP: dict[str, str] = {}
for canonical, aliases in HEADER_ALIASES.items():
    for alias in [canonical, *aliases]:
        key = _header_key(alias)
        if key and key not in ALIAS_LOOKUP:
            ALIAS_LOOKUP[key] = canonical


CANONICAL_TYPE_HINTS: dict[str, str] = {
    "Nomor Pelaporan": "report_number",
    "Tanggal": "date",
    "Bulan": "month",
    "Tahun": "year",
    "CUSTOMER_ID": "identity",
    "Nama Pelanggan": "person_name",
    "Nama Kontak": "person_name",
    "Nama Perusahaan": "company_name",
    "Kanwil": "category",
    "Kantor Cabang": "category",
    "Nomor Hp Pelanggan": "phone",
    "WhatsApp": "phone",
    "Email pelanggan": "email",
    "Nama Pemegang Polis": "person_name",
    "Nomor Polis": "policy_number",
    "POLICY_ID": "policy_number",
    "Produk Yang Digunakan": "category",
    "Jenis Layanan": "category",
    "Kategori": "category",
    "COB": "category",
    "Produk": "category",
    "Keterangan": "text",
    "Tindak Lanjut": "text",
    "Status": "category",
    "Saluran": "category",
    "Sifat Pelaporan": "category",
    "Agent": "person_name",
    "KTP_DEBITUR": "nik",
    "NPWP_DEBITUR": "npwp",
    "ALAMAT_DEBITUR": "text",
    "KELURAHAN_DEBITUR": "category",
    "KECAMATAN_DEBITUR": "category",
    "KOTA_DEBITUR": "category",
    "PROVINSI_DEBITUR": "category",
    "KOTA": "category",
    "PROVINSI": "category",
    "DEBITUR_ID": "identity",
    "BRANCH_ID": "identity",
    "POSTING_DATE": "date",
    "BANK_INDUK": "company_name",
    "ID_BANK_CABANG": "identity",
    "BANK_CABANG": "company_name",
    "ALAMAT_BANK_CABANG": "text",
    "KODE_POS": "identity",
    "KATEGORI_BANK": "category",
    "TIPE_DEBITUR": "category",
    "NAMA_DEBITUR": "person_name",
    "JENIS_KELAMIN_DEBITUR": "category",
    "TEMPAT_LAHIR_DEBITUR": "category",
    "TANGGAL_LAHIR_DEBITUR": "date",
    "KODE_POS_DEBITUR": "identity",
    "NEGARA_DEBITUR": "category",
    "REPORTING_DATE": "date",
}


@dataclass(frozen=True)
class CellResult:
    normalized: Any
    is_valid: bool | None
    reason: str
    issue_category: str
    changed: bool


@dataclass
class ETLConfig:
    input_file: str | Path
    output_file: str | Path
    input_filename: str = ""
    sheet_name: str | int = 0
    header: int | None = 0
    selected_columns: list[str] | None = None
    pipeline_mode: PipelineMode = "Generic"
    column_mapping: dict[str, str] = field(default_factory=dict)
    canonical_mapping: dict[str, str] = field(default_factory=dict)
    required_columns: list[str] = field(default_factory=list)
    duplicate_key_columns: list[str] = field(default_factory=list)
    placeholder_values: list[str] = field(default_factory=lambda: list(DEFAULT_PLACEHOLDERS))
    preserve_original_mode: Literal["all", "changed", "audit_only"] = "all"
    remove_empty_required: bool = False
    remove_invalid_phone: bool = False
    remove_invalid_email: bool = False
    remove_invalid_other: bool = False
    remove_duplicates: bool = False
    duplicate_keep: Literal["first", "last", "none"] = "first"
    create_audit: bool = True
    create_detail_reports: bool = True
    min_year: int = 2000
    max_year: int = field(default_factory=lambda: datetime.today().year)
    nik_lengths: list[int] = field(default_factory=lambda: [16])
    npwp_lengths: list[int] = field(default_factory=lambda: [15, 16])
    identity_lengths: list[int] = field(default_factory=list)
    policy_min_length: int = 3
    report_number_min_length: int = 3
    output_clean_sheet: str = "Data Bersih"
    output_sheets: list[str] = field(default_factory=lambda: list(DEFAULT_OUTPUT_SHEETS))
    progress_callback: Callable[[int, str], None] | None = None

    # Backward-compatible aliases. They no longer force hidden profile logic.
    drop_anomalous_phone: bool | None = None
    drop_missing_email_product: bool | None = None
    keep_original_columns: bool | None = None

    def __post_init__(self) -> None:
        if self.drop_anomalous_phone is not None:
            self.remove_invalid_phone = bool(self.drop_anomalous_phone)
        if self.drop_missing_email_product is not None:
            self.remove_empty_required = bool(self.drop_missing_email_product)
        if self.keep_original_columns is not None:
            self.preserve_original_mode = "all" if self.keep_original_columns else "audit_only"
        self.pipeline_mode = str(self.pipeline_mode).strip() or "Generic"  # type: ignore[assignment]
        self.selected_columns = list(self.selected_columns or [])
        self.required_columns = list(dict.fromkeys(self.required_columns))
        self.duplicate_key_columns = list(dict.fromkeys(self.duplicate_key_columns))
        self.placeholder_values = list(dict.fromkeys(self.placeholder_values or DEFAULT_PLACEHOLDERS))
        self.output_sheets = list(dict.fromkeys(self.output_sheets or DEFAULT_OUTPUT_SHEETS))
        self.pipeline_mode = "Generic"


# ============================================================
# PLACEHOLDER DAN MISSING VALUE
# ============================================================


def _placeholder_key(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.strip().upper()
    text = re.sub(r"[<>_./\\-]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_original_empty(value: Any) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def is_placeholder(value: Any, placeholders: Iterable[str] | None = None) -> bool:
    if is_original_empty(value):
        return False

    key = _placeholder_key(value)
    configured = {_placeholder_key(item) for item in (placeholders or DEFAULT_PLACEHOLDERS)}
    if key in configured:
        return True

    has_migration = any(token in key for token in ("MIGRATION", "MIGRASI", "MIGRATE"))
    has_address = any(token in key for token in ("ADDRESS", "ADDRES", "ALAMAT", "M ADDRESS"))
    has_empty = any(token in key for token in ("KOSONG", "NULL", "EMPTY", "NO DATA", "TIDAK ADA", "UNKNOWN"))
    return has_address and (has_migration or has_empty)


def is_missing_or_placeholder(value: Any, placeholders: Iterable[str] | None = None) -> bool:
    """Kompatibilitas fungsi lama: kosong asli atau placeholder."""
    return is_original_empty(value) or is_placeholder(value, placeholders)


def classify_source_value(value: Any, placeholders: Iterable[str]) -> str:
    if is_original_empty(value):
        return "EMPTY"
    if is_placeholder(value, placeholders):
        return "PLACEHOLDER"
    return "VALUE"


# ============================================================
# REKOMENDASI DAN VALIDASI KONFIGURASI
# ============================================================


def recommend_column_type(column_name: Any, pipeline_mode: str = "Generic") -> str:
    key = _header_key(column_name)
    canonical = ALIAS_LOOKUP.get(key)
    if canonical and canonical in CANONICAL_TYPE_HINTS:
        return CANONICAL_TYPE_HINTS[canonical]

    name = str(column_name).casefold()
    heuristic_rules = [
        (("whatsapp", "nomor hp", "no hp", "handphone", "telepon", "phone", "mobile", "wa"), "phone"),
        (("email", "e-mail"), "email"),
        (("tanggal", "date", "tgl"), "date"),
        (("bulan", "month"), "month"),
        (("tahun", "year"), "year"),
        (("npwp",), "npwp"),
        (("nik", "ktp"), "nik"),
        (("nomor polis", "no polis", "policy"), "policy_number"),
        (("nomor pelaporan", "no pelaporan", "ticket"), "report_number"),
        (("nama perusahaan", "company", "instansi", "bank"), "company_name"),
        (("nama", "responden", "kontak", "pic", "debitur"), "person_name"),
        (("jumlah", "nilai", "nominal", "amount", "total", "score"), "numeric"),
        (("status", "kategori", "produk", "cabang", "kanwil", "provinsi", "kota"), "category"),
    ]
    for needles, result in heuristic_rules:
        if any(needle in name for needle in needles):
            return result
    return "free" if pipeline_mode == "Custom" else "text"


def recommend_canonical_name(column_name: Any) -> str:
    return ALIAS_LOOKUP.get(_header_key(column_name), str(column_name).strip())


def build_mapping_recommendations(
    columns: Iterable[Any],
    pipeline_mode: str = "Generic",
) -> pd.DataFrame:
    """Membuat rekomendasi awal tipe kolom untuk satu pipeline Generic.

    Rekomendasi selalu ditampilkan kepada pengguna dan tidak menjadi keputusan
    final sampai mapping dikonfirmasi pada antarmuka.
    """
    rows: list[dict[str, Any]] = []
    for column in columns:
        source = str(column)
        canonical = recommend_canonical_name(source)
        column_type = recommend_column_type(source, "Generic")
        rows.append({
            "Analisis": True,
            "Kolom Sumber": source,
            "Nama Kanonik": canonical,
            "Rekomendasi Tipe": column_type,
            "Tipe Kolom": column_type,
        })
    return pd.DataFrame(rows)


def validate_etl_config(
    config: ETLConfig,
    available_columns: Iterable[Any],
) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []
    available = [str(column) for column in available_columns]
    available_set = set(available)

    if config.pipeline_mode != "Generic":
        errors.append("Pipeline hanya mendukung konfigurasi Generic.")
    invalid_output_sheets = [key for key in config.output_sheets if key not in REPORT_REGISTRY]
    if invalid_output_sheets:
        errors.append("Sheet output tidak dikenali: " + ", ".join(invalid_output_sheets))
    if not config.output_sheets:
        errors.append("Minimal satu sheet output harus dipilih.")
    if config.header is not None and config.header < 0:
        errors.append("Baris header tidak boleh negatif.")
    if not config.selected_columns:
        errors.append("Minimal satu kolom harus dipilih.")

    missing_selected = [column for column in config.selected_columns if column not in available_set]
    if missing_selected:
        errors.append("Kolom terpilih tidak tersedia: " + ", ".join(missing_selected))

    missing_mapping = [column for column in config.selected_columns if column not in config.column_mapping]
    if missing_mapping:
        errors.append("Tipe kolom belum ditentukan: " + ", ".join(missing_mapping))

    invalid_types = {
        column: column_type
        for column, column_type in config.column_mapping.items()
        if column in config.selected_columns and column_type not in COLUMN_TYPE_LABELS
    }
    if invalid_types:
        errors.append("Tipe kolom tidak valid: " + json.dumps(invalid_types, ensure_ascii=False))

    missing_required = [column for column in config.required_columns if column not in config.selected_columns]
    if missing_required:
        errors.append("Kolom wajib tidak termasuk kolom terpilih: " + ", ".join(missing_required))

    missing_duplicate_keys = [column for column in config.duplicate_key_columns if column not in config.selected_columns]
    if missing_duplicate_keys:
        errors.append("Kunci duplikat tidak termasuk kolom terpilih: " + ", ".join(missing_duplicate_keys))

    canonical_by_source = {
        source: (config.canonical_mapping.get(source) or source).strip()
        for source in config.selected_columns
    }
    empty_canonical = [source for source, canonical in canonical_by_source.items() if not canonical]
    if empty_canonical:
        errors.append("Nama kanonik kosong pada: " + ", ".join(empty_canonical))

    grouped: dict[str, list[str]] = {}
    for source, canonical in canonical_by_source.items():
        grouped.setdefault(canonical, []).append(source)
    for canonical, sources in grouped.items():
        if len(sources) <= 1:
            continue
        types = {config.column_mapping.get(source) for source in sources}
        if len(types) > 1:
            errors.append(
                f"Kolom yang digabung ke '{canonical}' memiliki tipe bertentangan: "
                + ", ".join(sources)
            )
        else:
            warnings.append(
                f"Kolom {', '.join(sources)} akan digabung ke '{canonical}' dengan prioritas dari kiri ke kanan."
            )

    if config.min_year > config.max_year:
        errors.append("Tahun minimum tidak boleh melebihi tahun maksimum.")
    if config.report_number_min_length < 1 or config.policy_min_length < 1:
        errors.append("Panjang minimum nomor harus minimal 1.")
    if config.preserve_original_mode not in {"all", "changed", "audit_only"}:
        errors.append("Mode penyimpanan nilai asli tidak valid.")
    if config.duplicate_keep not in {"first", "last", "none"}:
        errors.append("Perlakuan duplikat tidak valid.")
    if config.remove_duplicates and not config.duplicate_key_columns:
        errors.append("Penghapusan duplikat diaktifkan tetapi kunci duplikat belum dipilih.")
    if not config.create_audit and config.preserve_original_mode == "audit_only":
        errors.append("Mode audit_only membutuhkan sheet Audit Perubahan.")

    if config.pipeline_mode in {"Generic", "Custom"} and all(
        config.column_mapping.get(column) == "free" for column in config.selected_columns
    ):
        warnings.append("Semua kolom bertipe bebas; validity tidak akan dinilai.")

    return errors, warnings


# ============================================================
# EXTRACT DAN HARMONISASI HEADER
# ============================================================


def resolve_selected_columns(
    available_columns: Iterable[Any],
    selected_columns: list[str] | None,
) -> list[Any] | None:
    if selected_columns is None:
        return None
    if not selected_columns:
        raise ValueError("selected_columns tidak boleh berupa list kosong.")

    available_columns = list(available_columns)
    available_lookup = {_header_key(column): column for column in available_columns}
    resolved: list[Any] = []
    missing: list[str] = []

    for requested in selected_columns:
        requested_key = _header_key(requested)
        matched = available_lookup.get(requested_key)
        if matched is None:
            # Alias hanya membantu menemukan header yang memang dipilih pengguna.
            canonical = ALIAS_LOOKUP.get(requested_key)
            possible_names = [canonical, *HEADER_ALIASES.get(canonical, [])] if canonical else []
            for possible_name in possible_names:
                possible_key = _header_key(possible_name)
                if possible_key in available_lookup:
                    matched = available_lookup[possible_key]
                    break
        if matched is None:
            missing.append(str(requested))
        elif matched not in resolved:
            resolved.append(matched)

    if missing:
        raise KeyError(
            "Kolom tidak ditemukan: " + ", ".join(missing)
            + ". Kolom tersedia: " + ", ".join(map(str, available_columns))
        )
    return resolved


def read_preview_data(
    file_bytes: bytes,
    sheet_name: str | int,
    header_row: int,
    nrows: int = 20,
    usecols: list[Any] | None = None,
) -> pd.DataFrame:
    with io.BytesIO(file_bytes) as buffer:
        excel_file = pd.ExcelFile(buffer)
        return pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            header=header_row,
            nrows=nrows,
            usecols=usecols,
        )


def get_excel_sheet_names(file_bytes: bytes) -> list[str]:
    with io.BytesIO(file_bytes) as buffer:
        return pd.ExcelFile(buffer).sheet_names


def _read_csv_with_fallback(
    path: Path,
    *,
    header: int | None,
    nrows: int | None = None,
    usecols: list[Any] | None = None,
) -> pd.DataFrame:
    last_error: Exception | None = None
    for encoding in ["utf-8-sig", "utf-8", "cp1252", "latin-1"]:
        try:
            return pd.read_csv(
                path,
                header=header,
                nrows=nrows,
                usecols=usecols,
                encoding=encoding,
                sep=None,
                engine="python",
            )
        except (UnicodeDecodeError, pd.errors.ParserError) as exc:
            last_error = exc
    raise ValueError(f"CSV tidak dapat dibaca: {path}") from last_error


def extract_data(
    file_path: str | Path,
    sheet_name: str | int = 0,
    header: int | None = 0,
    selected_columns: list[str] | None = None,
    csv_separator: str | None = None,
    csv_encoding: str | None = None,
) -> pd.DataFrame:
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File tidak ditemukan: {path}")
    if not path.is_file():
        raise ValueError(f"Path sumber bukan file: {path}")

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        try:
            excel_file = pd.ExcelFile(path)
        except Exception as exc:
            raise ValueError(f"File Excel tidak dapat dibuka: {path}") from exc
        if isinstance(sheet_name, str) and sheet_name not in excel_file.sheet_names:
            raise ValueError(
                f"Sheet '{sheet_name}' tidak ditemukan. Sheet tersedia: {', '.join(excel_file.sheet_names)}"
            )
        preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=header, nrows=0)
        resolved = resolve_selected_columns(preview.columns, selected_columns)
        data = pd.read_excel(excel_file, sheet_name=sheet_name, header=header, usecols=resolved)
    elif suffix == ".csv":
        kwargs: dict[str, Any] = {}
        if csv_separator is not None:
            kwargs["sep"] = csv_separator
        if csv_encoding is not None:
            kwargs["encoding"] = csv_encoding
        if kwargs:
            preview = pd.read_csv(path, header=header, nrows=0, **kwargs)
            resolved = resolve_selected_columns(preview.columns, selected_columns)
            data = pd.read_csv(path, header=header, usecols=resolved, **kwargs)
        else:
            preview = _read_csv_with_fallback(path, header=header, nrows=0)
            resolved = resolve_selected_columns(preview.columns, selected_columns)
            data = _read_csv_with_fallback(path, header=header, usecols=resolved)
    else:
        raise ValueError(f"Format belum didukung: {suffix}. Gunakan Excel atau CSV.")

    if not isinstance(data, pd.DataFrame) or len(data.columns) == 0:
        raise ValueError("File tidak memiliki kolom data yang dapat dibaca.")
    if data.empty:
        raise ValueError("File tidak memiliki baris data.")
    if data.dropna(how="all").empty:
        raise ValueError("Seluruh baris data kosong.")
    return data


def harmonize_headers(
    df: pd.DataFrame,
    explicit_mapping: dict[str, str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Menerapkan mapping yang telah disetujui pengguna.

    Jika beberapa kolom dipetakan ke nama yang sama, nilai digabung dengan
    prioritas kiri ke kanan. Tidak ada auto-mapping tersembunyi.
    """
    mapping = explicit_mapping or {}
    groups: dict[str, list[str]] = {}
    for column in df.columns:
        source = str(column)
        target = mapping.get(source, source).strip() or source
        groups.setdefault(target, []).append(source)

    result = pd.DataFrame(index=df.index)
    mapping_rows: list[dict[str, Any]] = []
    for target, sources in groups.items():
        combined = df[sources[0]].copy()
        for source in sources[1:]:
            combined = combined.combine_first(df[source])
        result[target] = combined
        for source in sources:
            mapping_rows.append({
                "HEADER_ASLI": source,
                "HEADER_KANONIK": target,
                "DIGABUNG_DENGAN_KOLOM_LAIN": len(sources) > 1,
            })
    return result, pd.DataFrame(mapping_rows)


def add_stable_row_id(df: pd.DataFrame, header: int | None = 0) -> pd.DataFrame:
    work = df.copy().reset_index(drop=True)
    source_start = 1 if header is None else header + 2
    work.insert(0, "SOURCE_ROW_NUMBER", np.arange(source_start, source_start + len(work)))
    work.insert(0, "ROW_ID", [f"ROW-{position:07d}" for position in range(1, len(work) + 1)])
    return work


# ============================================================
# NORMALISASI DETAIL
# ============================================================


def _safe_equal(left: Any, right: Any) -> bool:
    if pd.isna(left) and pd.isna(right):
        return True
    try:
        return str(left) == str(right)
    except Exception:
        return False


def normalize_phone(value: Any) -> Any:
    """Fungsi lama dipertahankan; mengembalikan nilai normalisasi saja."""
    return normalize_phone_detailed(value).normalized


def normalize_phone_detailed(value: Any) -> CellResult:
    if is_original_empty(value):
        return CellResult(pd.NA, None, "Nilai kosong asli", "EMPTY", False)

    if isinstance(value, (int, float, np.integer, np.floating)):
        try:
            text = format(Decimal(str(value)), "f")
        except InvalidOperation:
            text = str(value)
    else:
        text = str(value)
    original_text = text.strip()

    cleaned = re.sub(
        r"(?i)\b(FAX|FACSIMILE|TELP|TELEPON|TEL|PHONE|HP|HANDPHONE|MOBILE|WA|WHATSAPP)\b\s*[:.]?",
        " ",
        original_text,
    )
    cleaned = re.sub(r"\.0+$", "", cleaned)
    candidates = re.findall(r"\+?\d[\d\s().-]{6,}\d|\+?\d{8,15}", cleaned)

    normalized: list[str] = []
    invalid_candidates: list[str] = []
    remainder = cleaned
    for candidate in candidates:
        remainder = remainder.replace(candidate, " ", 1)
        number = re.sub(r"\D", "", candidate)
        if not number or set(number) == {"0"}:
            invalid_candidates.append(candidate)
            continue
        if number.startswith("0062"):
            number = "62" + number[4:]
        elif number.startswith("620"):
            number = "62" + number[3:]
        elif number.startswith("62"):
            pass
        elif number.startswith("0"):
            number = "62" + number[1:]
        elif number.startswith("8"):
            number = "62" + number
        elif number.startswith(("2", "3", "4", "5", "6", "7", "9")):
            number = "62" + number
        else:
            invalid_candidates.append(candidate)
            continue
        final = "+" + number
        if re.fullmatch(r"\+62\d{7,13}", final):
            normalized.append(final)
        else:
            invalid_candidates.append(candidate)

    normalized = list(dict.fromkeys(normalized))
    remainder = re.sub(r"[\s,;/|\\]+", "", remainder)
    normalized_value = "; ".join(normalized) if normalized else pd.NA

    if not normalized:
        digits = re.sub(r"\D", "", original_text)
        if not digits:
            reason = "Tidak mengandung angka"
        elif len(digits) < 9:
            reason = "Nomor terlalu pendek"
        elif len(digits) > 30:
            reason = "Nomor terlalu panjang atau beberapa nomor tidak terpisah dengan benar"
        elif not digits.startswith(("0", "8", "62", "0062")):
            reason = "Awalan nomor tidak dikenali"
        else:
            reason = "Nomor gagal dinormalisasi"
        return CellResult(pd.NA, False, reason, "NORMALIZATION_FAILED", True)

    reasons: list[str] = []
    if invalid_candidates:
        reasons.append("Sebagian kandidat nomor tidak valid")
    if remainder:
        reasons.append(f"Terdapat teks sisa: {remainder[:80]}")
    is_valid = not reasons
    reason = "; ".join(reasons) if reasons else "Format nomor valid"
    return CellResult(
        normalized_value,
        is_valid,
        reason,
        "VALID" if is_valid else "INVALID",
        not _safe_equal(original_text, normalized_value),
    )


def normalize_email(value: Any) -> Any:
    return normalize_email_detailed(value).normalized


def normalize_email_detailed(value: Any) -> CellResult:
    if is_original_empty(value):
        return CellResult(pd.NA, None, "Nilai kosong asli", "EMPTY", False)
    text = str(value).strip().lower()
    email_pattern = r"[a-z0-9._%+-]+@[a-z0-9-]+(?:\.[a-z0-9-]+)+"
    emails = re.findall(email_pattern, text, flags=re.IGNORECASE)
    emails = list(dict.fromkeys(email.lower() for email in emails))
    remainder = re.sub(email_pattern, "", text, flags=re.IGNORECASE)
    remainder = re.sub(r"[\s,;/|\\<>()\[\]{}'`:\"]+", "", remainder)

    if not emails:
        return CellResult(pd.NA, False, "Tidak ditemukan alamat email valid", "NORMALIZATION_FAILED", True)

    normalized = "; ".join(emails)
    is_valid = remainder == ""
    reason = "Format email valid" if is_valid else f"Terdapat teks sisa bukan email: {remainder[:100]}"
    return CellResult(
        normalized,
        is_valid,
        reason,
        "VALID" if is_valid else "INVALID",
        not _safe_equal(text, normalized),
    )


INDONESIAN_MONTHS = {
    "JANUARI": "01", "FEBRUARI": "02", "MARET": "03", "APRIL": "04",
    "MEI": "05", "JUNI": "06", "JULI": "07", "AGUSTUS": "08",
    "SEPTEMBER": "09", "OKTOBER": "10", "NOVEMBER": "11", "DESEMBER": "12",
}


def normalize_date(value: Any) -> Any:
    return normalize_date_detailed(value, 2000, datetime.today().year).normalized


def normalize_date_detailed(value: Any, min_year: int, max_year: int) -> CellResult:
    if is_original_empty(value):
        return CellResult(pd.NA, None, "Nilai kosong asli", "EMPTY", False)

    parsed: pd.Timestamp | None = None
    original = value
    if isinstance(value, (datetime, date, pd.Timestamp)):
        parsed = pd.Timestamp(value)
    else:
        text = re.sub(r"\s+", " ", str(value).strip()).upper()
        for month_name, month_number in INDONESIAN_MONTHS.items():
            text = re.sub(rf"\b{month_name}\b", month_number, text)
        formats = [
            "%d %m %Y", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%m-%d-%Y",
            "%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S",
            "%d %B %Y", "%d %b %Y", "%B %d %Y", "%b %d %Y",
        ]
        for fmt in formats:
            try:
                parsed = pd.Timestamp(datetime.strptime(text, fmt))
                break
            except ValueError:
                continue
        if parsed is None:
            candidate = pd.to_datetime(text, errors="coerce", dayfirst=True)
            if pd.notna(candidate):
                parsed = pd.Timestamp(candidate)

    if parsed is None or pd.isna(parsed):
        return CellResult(pd.NA, False, "Tanggal tidak dapat diparsing", "NORMALIZATION_FAILED", True)

    normalized = parsed.strftime("%m/%d/%Y")
    valid_range = min_year <= parsed.year <= max_year
    reason = "Format dan rentang tanggal valid" if valid_range else f"Tahun harus berada pada rentang {min_year}-{max_year}"
    return CellResult(
        normalized,
        valid_range,
        reason,
        "VALID" if valid_range else "INVALID",
        not _safe_equal(original, normalized),
    )


def normalize_id(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .str.strip()
        .str.replace(r"(?i)^(NPWP|IDNO|KTP|NIK)\s*[-:/]?\s*", "", regex=True)
        .str.replace(r"\.0$", "", regex=True)
        .str.replace(r"\D", "", regex=True)
        .replace("", pd.NA)
    )


def normalize_identity_detailed(
    value: Any,
    allowed_lengths: Iterable[int] | None,
    labels: Iterable[str],
    identity_name: str,
) -> CellResult:
    if is_original_empty(value):
        return CellResult(pd.NA, None, "Nilai kosong asli", "EMPTY", False)
    text = str(value).strip()
    labels_pattern = "|".join(re.escape(label) for label in labels)
    cleaned = re.sub(rf"(?i)^({labels_pattern})\s*[-:/]?\s*", "", text)
    cleaned = re.sub(r"\.0$", "", cleaned)
    digits = re.sub(r"\D", "", cleaned)
    if not digits:
        return CellResult(pd.NA, False, f"{identity_name} tidak mengandung digit", "NORMALIZATION_FAILED", True)

    lengths = sorted(set(int(length) for length in (allowed_lengths or []) if int(length) > 0))
    is_valid = not lengths or len(digits) in lengths
    if is_valid:
        reason = f"Format {identity_name} valid berdasarkan pola dan panjang"
    else:
        reason = f"Panjang {identity_name} {len(digits)} digit; panjang yang diterima: {lengths}"
    return CellResult(
        digits,
        is_valid,
        reason,
        "VALID" if is_valid else "INVALID",
        not _safe_equal(text, digits),
    )




LEGAL_ENTITY_TOKENS = {
    "PT", "CV", "UD", "PD", "PERUM", "PERSERO", "KOPERASI",
    "YAYASAN", "FIRMA", "FA", "TBK",
}


def normalize_company_name(value: Any) -> Any:
    """Menormalkan penulisan nama perusahaan tanpa menebak badan hukum.

    PT/CV hanya dipertahankan jika memang tersedia pada nilai sumber.
    Fungsi tidak menambahkan badan hukum yang tidak diketahui.
    """
    if is_original_empty(value):
        return pd.NA
    text = unicodedata.normalize("NFKC", str(value))
    text = re.sub(r"[\t\r\n]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return pd.NA

    replacements = [
        (r"(?i)\bP\.?\s*T\.?\b", "PT"),
        (r"(?i)\bC\.?\s*V\.?\b", "CV"),
        (r"(?i)\bU\.?\s*D\.?\b", "UD"),
        (r"(?i)\bT\.?\s*B\.?\s*K\.?\b", "TBK"),
    ]
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)

    words: list[str] = []
    for raw_word in text.split():
        word = raw_word.strip(".,;:")
        upper = word.upper()
        if upper in LEGAL_ENTITY_TOKENS:
            words.append(upper)
        elif word.isupper() and len(word) <= 4:
            words.append(word)
        else:
            words.append(word.title())
    return " ".join(words) if words else pd.NA


def company_match_key(value: Any) -> str:
    """Kunci toleran untuk membandingkan nama perusahaan.

    Badan hukum dihapus hanya dari kunci pencocokan, bukan dari nilai utama.
    """
    normalized = normalize_company_name(value)
    if pd.isna(normalized):
        return ""
    text = str(normalized).casefold()
    text = re.sub(r"^(?:pt|cv|ud|pd|firma|fa|koperasi|yayasan|perum|persero)\s+", "", text)
    text = re.sub(r"\s+(?:tbk|persero)$", "", text)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_person_name(value: Any) -> Any:
    if is_original_empty(value):
        return pd.NA
    text = unicodedata.normalize("NFKC", str(value))
    text = re.sub(r"[\t\r\n]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return pd.NA
    words: list[str] = []
    for raw_word in text.split(" "):
        words.append(raw_word.title())
    normalized = " ".join(words).strip()
    return normalized if normalized else pd.NA


def person_name_match_key(value: Any) -> str:
    normalized = normalize_person_name(value)
    if pd.isna(normalized):
        return ""
    text = str(normalized).casefold()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _normalize_simple_text(value: Any, kind: str) -> CellResult:
    if is_original_empty(value):
        return CellResult(pd.NA, None, "Nilai kosong asli", "EMPTY", False)
    original = str(value)
    normalized = re.sub(r"\s+", " ", original.strip())
    if kind in {"person_name", "company_name"}:
        valid = len(normalized) >= 2 and re.search(r"[A-Za-zÀ-ÿ]", normalized) is not None
        reason = "Nama valid" if valid else "Nama harus memiliki minimal 2 karakter dan mengandung huruf"
    elif kind == "category":
        valid = len(normalized) >= 1 and re.search(r"[A-Za-z0-9]", normalized) is not None
        reason = "Kategori valid" if valid else "Kategori tidak mengandung nilai bermakna"
    else:
        valid = len(normalized) >= 2 and re.search(r"[A-Za-z0-9]", normalized) is not None
        reason = "Teks valid" if valid else "Teks harus memiliki minimal 2 karakter alfanumerik"
    return CellResult(
        normalized,
        valid,
        reason,
        "VALID" if valid else "INVALID",
        original != normalized,
    )


def _normalize_month(value: Any) -> CellResult:
    if is_original_empty(value):
        return CellResult(pd.NA, None, "Nilai kosong asli", "EMPTY", False)
    text = str(value).strip().upper()
    month_names = {
        "JAN": 1, "JANUARY": 1, "JANUARI": 1,
        "FEB": 2, "FEBRUARY": 2, "FEBRUARI": 2,
        "MAR": 3, "MARCH": 3, "MARET": 3,
        "APR": 4, "APRIL": 4,
        "MAY": 5, "MEI": 5,
        "JUN": 6, "JUNE": 6, "JUNI": 6,
        "JUL": 7, "JULY": 7, "JULI": 7,
        "AUG": 8, "AUGUST": 8, "AGUSTUS": 8,
        "SEP": 9, "SEPT": 9, "SEPTEMBER": 9,
        "OCT": 10, "OCTOBER": 10, "OKTOBER": 10,
        "NOV": 11, "NOVEMBER": 11,
        "DEC": 12, "DECEMBER": 12, "DESEMBER": 12,
    }
    month: int | None = None
    try:
        numeric = float(text)
        if numeric.is_integer() and 1 <= int(numeric) <= 12:
            month = int(numeric)
    except ValueError:
        month = month_names.get(text)
    if month is None:
        return CellResult(pd.NA, False, "Bulan tidak berada pada rentang 1-12 atau nama bulan tidak dikenali", "NORMALIZATION_FAILED", True)
    return CellResult(month, True, "Bulan valid", "VALID", not _safe_equal(value, month))


def _normalize_year(value: Any, min_year: int, max_year: int) -> CellResult:
    if is_original_empty(value):
        return CellResult(pd.NA, None, "Nilai kosong asli", "EMPTY", False)
    try:
        year = int(float(str(value).strip()))
    except (ValueError, TypeError):
        return CellResult(pd.NA, False, "Tahun bukan bilangan bulat", "NORMALIZATION_FAILED", True)
    valid = min_year <= year <= max_year
    reason = "Tahun valid" if valid else f"Tahun harus berada pada rentang {min_year}-{max_year}"
    return CellResult(year, valid, reason, "VALID" if valid else "INVALID", not _safe_equal(value, year))


def _normalize_numeric(value: Any) -> CellResult:
    if is_original_empty(value):
        return CellResult(pd.NA, None, "Nilai kosong asli", "EMPTY", False)
    text = str(value).strip().replace(" ", "")
    # Bentuk Indonesia sederhana: 1.234,56 -> 1234.56
    if re.fullmatch(r"[-+]?\d{1,3}(?:\.\d{3})+(?:,\d+)?", text):
        text = text.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"[-+]?\d+(?:,\d+)", text):
        text = text.replace(",", ".")
    try:
        number = Decimal(text)
    except InvalidOperation:
        return CellResult(pd.NA, False, "Nilai bukan angka yang dapat diparsing", "NORMALIZATION_FAILED", True)
    normalized: int | float = int(number) if number == number.to_integral_value() else float(number)
    return CellResult(normalized, True, "Nilai numerik valid", "VALID", not _safe_equal(value, normalized))


def _normalize_policy_or_report(value: Any, min_length: int, label: str) -> CellResult:
    if is_original_empty(value):
        return CellResult(pd.NA, None, "Nilai kosong asli", "EMPTY", False)
    original = str(value)
    normalized = re.sub(r"\s+", "", original.strip())
    allowed = re.fullmatch(r"[A-Za-z0-9./_-]+", normalized) is not None
    valid = len(normalized) >= min_length and allowed and re.search(r"[A-Za-z0-9]", normalized) is not None
    if valid:
        reason = f"{label} valid"
    elif len(normalized) < min_length:
        reason = f"{label} kurang dari {min_length} karakter"
    else:
        reason = f"{label} mengandung karakter yang tidak diperbolehkan"
    return CellResult(normalized, valid, reason, "VALID" if valid else "INVALID", original != normalized)


def normalize_cell(
    value: Any,
    column_type: str,
    config: ETLConfig,
) -> CellResult:
    source_class = classify_source_value(value, config.placeholder_values)
    if source_class == "EMPTY":
        return CellResult(pd.NA, None, "Nilai kosong asli", "EMPTY", False)
    if source_class == "PLACEHOLDER":
        return CellResult(pd.NA, None, f"Placeholder terdeteksi: {value}", "PLACEHOLDER", True)

    if column_type == "free":
        return CellResult(value, None, "Kolom tidak dinormalisasi dan tidak dinilai", "NOT_ASSESSED", False)
    if column_type == "company_name":
        normalized_company = normalize_company_name(value)
        if pd.isna(normalized_company):
            return CellResult(pd.NA, False, "Nama perusahaan tidak memiliki nilai bermakna", "NORMALIZATION_FAILED", True)
        valid = len(str(normalized_company)) >= 2 and re.search(r"[A-Za-zÀ-ÿ]", str(normalized_company)) is not None
        return CellResult(
            normalized_company,
            valid,
            "Nama perusahaan valid" if valid else "Nama perusahaan harus mengandung huruf",
            "VALID" if valid else "INVALID",
            not _safe_equal(value, normalized_company),
        )
    if column_type in {"text", "person_name", "category"}:
        return _normalize_simple_text(value, column_type)
    if column_type == "phone":
        return normalize_phone_detailed(value)
    if column_type == "email":
        return normalize_email_detailed(value)
    if column_type == "date":
        return normalize_date_detailed(value, config.min_year, config.max_year)
    if column_type == "month":
        return _normalize_month(value)
    if column_type == "year":
        return _normalize_year(value, config.min_year, config.max_year)
    if column_type == "numeric":
        return _normalize_numeric(value)
    if column_type == "nik":
        return normalize_identity_detailed(value, config.nik_lengths, ["NIK", "KTP", "IDNO"], "NIK")
    if column_type == "npwp":
        return normalize_identity_detailed(value, config.npwp_lengths, ["NPWP"], "NPWP")
    if column_type == "identity":
        return normalize_identity_detailed(value, config.identity_lengths, ["NIK", "KTP", "NPWP", "IDNO"], "nomor identitas")
    if column_type == "policy_number":
        return _normalize_policy_or_report(value, config.policy_min_length, "Nomor polis")
    if column_type == "report_number":
        return _normalize_policy_or_report(value, config.report_number_min_length, "Nomor pelaporan")
    raise ValueError(f"Tipe kolom tidak didukung: {column_type}")


def validate_value(column: str, value: Any) -> bool | float:
    """Aturan validitas lama yang dipertahankan sebagai sumber kebenaran.

    True  = terisi dan valid
    False = terisi tetapi tidak valid
    NaN   = kosong/placeholder dan tidak dinilai sebagai invalid
    """
    if is_missing_or_placeholder(value):
        return np.nan

    text = str(value).strip()

    date_columns = {
        "Tanggal", "POSTING_DATE", "REPORTING_DATE", "TANGGAL_LAHIR_DEBITUR"
    }
    if column in date_columns:
        normalized = normalize_date(value)
        if pd.isna(normalized):
            return False
        try:
            parsed = datetime.strptime(str(normalized), "%m/%d/%Y")
        except (ValueError, TypeError):
            return False

        today = datetime.today()
        if column == "TANGGAL_LAHIR_DEBITUR":
            return datetime(1900, 1, 1) <= parsed <= today
        return datetime(2000, 1, 1) <= parsed <= today

    if column == "Nomor Pelaporan":
        return (
            len(text) >= 3
            and re.search(r"[A-Za-z0-9]", text) is not None
            and re.fullmatch(r"[A-Za-z0-9./_-]+", text) is not None
        )

    if column == "Bulan":
        month_names = {
            "JAN", "JANUARY", "JANUARI", "FEB", "FEBRUARY", "FEBRUARI",
            "MAR", "MARCH", "MARET", "APR", "APRIL", "MAY", "MEI",
            "JUN", "JUNE", "JUNI", "JUL", "JULY", "JULI",
            "AUG", "AUGUST", "AGUSTUS", "SEP", "SEPT", "SEPTEMBER",
            "OCT", "OCTOBER", "OKTOBER", "NOV", "NOVEMBER",
            "DEC", "DECEMBER", "DESEMBER",
        }
        try:
            month = float(text)
            if month.is_integer():
                return 1 <= int(month) <= 12
        except ValueError:
            pass
        return text.upper() in month_names

    if column == "Tahun":
        try:
            year = int(float(text))
            return 2000 <= year <= datetime.today().year
        except (ValueError, TypeError):
            return False

    if column in {
        "Nama Pelanggan", "Nama Kontak", "NAMA_DEBITUR",
        "Nama Perusahaan", "BANK_INDUK", "BANK_CABANG",
        "TEMPAT_LAHIR_DEBITUR", "KELURAHAN_DEBITUR",
        "KECAMATAN_DEBITUR", "KOTA_DEBITUR", "PROVINSI_DEBITUR",
        "KOTA", "PROVINSI", "NEGARA_DEBITUR"
    }:
        return (
            len(text) >= 2
            and re.search(r"[A-Za-zÀ-ÿ]", text) is not None
            and re.fullmatch(r"[^\n\r\t]+", text) is not None
        )

    if column in {"ALAMAT_DEBITUR", "ALAMAT_BANK_CABANG"}:
        compact = re.sub(r"\s+", " ", text).strip()
        return (
            len(compact) >= 5
            and re.search(r"[A-Za-zÀ-ÿ]", compact) is not None
            and re.search(r"[A-Za-z0-9]", compact) is not None
            and not is_missing_or_placeholder(compact)
        )

    if column in {"KODE_POS", "KODE_POS_DEBITUR"}:
        digits = re.sub(r"\D", "", text)
        return re.fullmatch(r"\d{5}", digits) is not None

    if column == "JENIS_KELAMIN_DEBITUR":
        key = _placeholder_key(text)
        return key in {
            "L", "P", "LAKI LAKI", "LAKI", "PRIA", "MALE", "M",
            "PEREMPUAN", "WANITA", "FEMALE", "F"
        }

    if column in {"Nomor Hp Pelanggan", "WhatsApp"}:
        numbers = [item.strip() for item in text.split(";") if item.strip()]
        return bool(numbers) and all(
            re.fullmatch(r"\+62\d{7,13}", number) is not None
            for number in numbers
        )

    if column == "Email pelanggan":
        email_pattern = r"[a-z0-9._%+-]+@[a-z0-9-]+(?:\.[a-z0-9-]+)+"
        emails = re.findall(email_pattern, text.lower(), flags=re.IGNORECASE)
        if not emails:
            return False
        emails = list(dict.fromkeys(emails))
        remainder = re.sub(email_pattern, "", text.lower(), flags=re.IGNORECASE)
        remainder = re.sub(r"[\s,;/|\\<>()\[\]{}'`:]+", "", remainder)
        return not remainder and all(
            re.fullmatch(email_pattern, email, flags=re.IGNORECASE)
            for email in emails
        )

    if column in {"Nomor Polis", "POLICY_ID"}:
        return (
            len(text) >= 3
            and re.search(r"[A-Za-z0-9]", text) is not None
            and re.fullmatch(r"[A-Za-z0-9./_-]+", text) is not None
        )

    if column in {"CUSTOMER_ID", "DEBITUR_ID", "BRANCH_ID", "ID_BANK_CABANG"}:
        return (
            len(text) >= 1
            and re.search(r"[A-Za-z0-9]", text) is not None
            and re.fullmatch(r"[A-Za-z0-9./_-]+", text) is not None
        )

    if column == "KTP_DEBITUR":
        return re.fullmatch(r"\d{16}", text) is not None

    if column == "NPWP_DEBITUR":
        return re.fullmatch(r"\d{15}|\d{16}", text) is not None

    if column in {"KATEGORI_BANK", "TIPE_DEBITUR"}:
        return len(text) >= 1 and re.search(r"[A-Za-z0-9]", text) is not None

    return len(text) >= 2 and re.search(r"[A-Za-z0-9]", text) is not None


LEGACY_VALIDATION_COLUMNS = {
    "Nomor Pelaporan", "Tanggal", "Bulan", "Tahun", "Nama Pelanggan",
    "Nomor Hp Pelanggan", "Email pelanggan", "Nama Pemegang Polis",
    "Nomor Polis", "Nama Perusahaan", "Jenis Layanan", "Kategori", "COB",
    "Produk", "Keterangan", "Tindak Lanjut", "Status", "Saluran",
    "Sifat Pelaporan", "Agent", "WhatsApp", "Nama Kontak",
    "POLICY_ID", "CUSTOMER_ID", "DEBITUR_ID", "BRANCH_ID",
    "POSTING_DATE", "REPORTING_DATE", "BANK_INDUK", "ID_BANK_CABANG",
    "BANK_CABANG", "ALAMAT_BANK_CABANG", "KODE_POS", "KOTA",
    "PROVINSI", "KATEGORI_BANK", "TIPE_DEBITUR", "NAMA_DEBITUR",
    "JENIS_KELAMIN_DEBITUR", "TEMPAT_LAHIR_DEBITUR",
    "TANGGAL_LAHIR_DEBITUR", "KTP_DEBITUR", "NPWP_DEBITUR",
    "ALAMAT_DEBITUR", "KODE_POS_DEBITUR", "KELURAHAN_DEBITUR",
    "KECAMATAN_DEBITUR", "KOTA_DEBITUR", "PROVINSI_DEBITUR",
    "NEGARA_DEBITUR",
}


def _apply_legacy_validation(
    column: str,
    original_value: Any,
    result: CellResult,
) -> CellResult:
    """Memastikan kolom lama tetap memakai aturan invalid yang telah ditetapkan.

    Normalisasi baru dan audit tetap digunakan. Untuk kolom kanonik yang sudah
    memiliki aturan bisnis lama, status valid/invalid ditentukan oleh
    ``validate_value`` terhadap nilai hasil normalisasi. Nilai kosong dan
    placeholder tetap tidak dinilai.
    """
    if column not in LEGACY_VALIDATION_COLUMNS:
        return result
    if result.issue_category in {"EMPTY", "PLACEHOLDER"}:
        return result

    value_to_validate = result.normalized
    if pd.isna(value_to_validate):
        legacy_valid = False
    else:
        legacy_valid = validate_value(column, value_to_validate)

    if pd.isna(legacy_valid):
        return CellResult(
            result.normalized,
            None,
            result.reason,
            result.issue_category,
            result.changed,
        )

    is_valid = bool(legacy_valid)
    if is_valid:
        reason = result.reason if result.is_valid is True else "Memenuhi aturan validasi lama"
        category = "VALID"
    else:
        reason = f"Tidak memenuhi aturan validasi lama untuk kolom {column}"
        if result.reason and result.reason not in {"Format valid", reason}:
            reason = f"{reason}; {result.reason}"
        category = "NORMALIZATION_FAILED" if pd.isna(result.normalized) else "INVALID"

    return CellResult(
        result.normalized,
        is_valid,
        reason,
        category,
        result.changed,
    )


# ============================================================
# TRANSFORMASI, DUPLIKAT, QUALITY REPORT, CLEANING
# ============================================================


def _effective_config_mappings(
    config: ETLConfig,
    header_mapping: pd.DataFrame,
) -> tuple[dict[str, str], list[str], list[str]]:
    source_to_target = dict(zip(header_mapping["HEADER_ASLI"], header_mapping["HEADER_KANONIK"]))
    type_mapping: dict[str, str] = {}
    for source in config.selected_columns:
        target = source_to_target.get(source, source)
        source_type = config.column_mapping[source]
        if target in type_mapping and type_mapping[target] != source_type:
            raise ValueError(f"Tipe bertentangan untuk kolom gabungan '{target}'.")
        type_mapping[target] = source_type
    required = list(dict.fromkeys(source_to_target.get(column, column) for column in config.required_columns))
    duplicate_keys = list(dict.fromkeys(source_to_target.get(column, column) for column in config.duplicate_key_columns))
    return type_mapping, required, duplicate_keys


def transform_data(
    df: pd.DataFrame,
    keep_original_columns: bool = True,
    *,
    config: ETLConfig | None = None,
    type_mapping: dict[str, str] | None = None,
) -> tuple[pd.DataFrame, dict[str, pd.DataFrame]]:
    """Normalisasi kolom berdasarkan mapping eksplisit.

    Parameter lama ``keep_original_columns`` tetap diterima. ``config`` dan
    ``type_mapping`` digunakan oleh versi baru.
    """
    if config is None:
        config = ETLConfig(
            input_file=".",
            output_file=".",
            selected_columns=[column for column in df.columns if column not in {"ROW_ID", "SOURCE_ROW_NUMBER"}],
            column_mapping={
                column: recommend_column_type(column)
                for column in df.columns if column not in {"ROW_ID", "SOURCE_ROW_NUMBER"}
            },
            preserve_original_mode="all" if keep_original_columns else "audit_only",
        )
    type_mapping = type_mapping or {
        column: config.column_mapping.get(column, "free")
        for column in df.columns if column not in {"ROW_ID", "SOURCE_ROW_NUMBER"}
    }

    normalized = df.copy()
    audit_rows: list[dict[str, Any]] = []
    cell_rows: list[dict[str, Any]] = []
    clean_values = df.copy()

    for column, column_type in type_mapping.items():
        if column not in df.columns:
            continue
        results = [
            _apply_legacy_validation(
                column,
                value,
                normalize_cell(value, column_type, config),
            )
            for value in df[column].tolist()
        ]
        normalized_values = [result.normalized for result in results]
        valid_values = [result.is_valid for result in results]
        reasons = [result.reason for result in results]
        issues = [result.issue_category for result in results]
        changed_values = [result.changed for result in results]

        if config.preserve_original_mode == "all":
            normalized[f"{column}_original"] = df[column]
        elif config.preserve_original_mode == "changed":
            normalized[f"{column}_original"] = [
                original if changed else pd.NA
                for original, changed in zip(df[column].tolist(), changed_values)
            ]
        normalized[f"{column}_normalized"] = normalized_values
        normalized[f"{column}_is_valid"] = pd.array(valid_values, dtype="boolean")
        normalized[f"{column}_validation_reason"] = reasons
        normalized[f"{column}_issue_category"] = issues

        # Data bersih memakai nilai normalisasi jika tersedia. Nilai invalid yang
        # gagal dinormalisasi tetap dipertahankan agar tidak hilang diam-diam.
        final_values: list[Any] = []
        for original, result in zip(df[column].tolist(), results):
            final_values.append(original if pd.isna(result.normalized) and result.issue_category in {"INVALID", "NORMALIZATION_FAILED"} else result.normalized)
        clean_values[column] = final_values
        if config.preserve_original_mode == "all":
            clean_values[f"{column}_original"] = df[column]
        elif config.preserve_original_mode == "changed":
            clean_values[f"{column}_original"] = [
                original if changed else pd.NA
                for original, changed in zip(df[column].tolist(), changed_values)
            ]

        for position, result in enumerate(results):
            record = {
                "ROW_ID": df.iloc[position]["ROW_ID"],
                "SOURCE_ROW_NUMBER": df.iloc[position]["SOURCE_ROW_NUMBER"],
                "COLUMN": column,
                "COLUMN_TYPE": column_type,
                "ORIGINAL_VALUE": df.iloc[position][column],
                "NORMALIZED_VALUE": result.normalized,
                "IS_VALID": result.is_valid,
                "ISSUE_CATEGORY": result.issue_category,
                "VALIDATION_REASON": result.reason,
                "CHANGED": result.changed,
            }
            cell_rows.append(record)
            if config.create_audit and (
                result.changed
                or result.issue_category in {"EMPTY", "PLACEHOLDER", "INVALID", "NORMALIZATION_FAILED"}
            ):
                audit_rows.append({
                    "ROW_ID": record["ROW_ID"],
                    "SOURCE_ROW_NUMBER": record["SOURCE_ROW_NUMBER"],
                    "COLUMN": column,
                    "ORIGINAL_VALUE": record["ORIGINAL_VALUE"],
                    "NORMALIZED_VALUE": record["NORMALIZED_VALUE"],
                    "CHANGE_TYPE": _change_type(result),
                    "VALIDATION_STATUS": _validation_status(result.is_valid),
                    "REASON": result.reason,
                })

    reports = {
        "Cell Assessment": pd.DataFrame(cell_rows, columns=[
            "ROW_ID", "SOURCE_ROW_NUMBER", "COLUMN", "COLUMN_TYPE",
            "ORIGINAL_VALUE", "NORMALIZED_VALUE", "IS_VALID",
            "ISSUE_CATEGORY", "VALIDATION_REASON", "CHANGED",
        ]),
        "Audit Perubahan": pd.DataFrame(audit_rows, columns=[
            "ROW_ID", "SOURCE_ROW_NUMBER", "COLUMN", "ORIGINAL_VALUE",
            "NORMALIZED_VALUE", "CHANGE_TYPE", "VALIDATION_STATUS", "REASON",
        ]),
        "Clean Values": clean_values,
    }
    return normalized, reports


def _change_type(result: CellResult) -> str:
    if result.issue_category == "EMPTY":
        return "EMPTY_ORIGINAL"
    if result.issue_category == "PLACEHOLDER":
        return "PLACEHOLDER_TO_EMPTY"
    if result.issue_category == "NORMALIZATION_FAILED":
        return "NORMALIZATION_FAILED"
    if result.issue_category == "INVALID":
        return "NORMALIZED_BUT_INVALID"
    if result.changed:
        return "NORMALIZED"
    return "UNCHANGED"


def _validation_status(value: bool | None) -> str:
    if value is True:
        return "VALID"
    if value is False:
        return "INVALID"
    return "NOT_ASSESSED"


def build_duplicate_report(
    clean_values: pd.DataFrame,
    duplicate_key_columns: list[str],
) -> pd.DataFrame:
    """Mendeteksi duplikat dengan logika OR untuk setiap kolom terpilih.

    Setiap kolom kunci diperiksa secara independen. Sebuah baris masuk laporan
    apabila nilainya berulang pada minimal satu kolom yang dipilih. Jika satu
    baris berulang pada beberapa kolom, baris tersebut muncul sekali untuk
    setiap aturan sehingga kolom dan nilai pemicu tetap dapat ditelusuri.
    """
    columns = [
        "ROW_ID", "SOURCE_ROW_NUMBER", "COMPARISON_COLUMN",
        "COMPARISON_VALUE", "NORMALIZED_COMPARISON_VALUE",
        "DUPLICATE_GROUP", "GROUP_SIZE", "STATUS", "REASON",
    ]
    if not duplicate_key_columns:
        return pd.DataFrame(columns=columns)

    missing = [column for column in duplicate_key_columns if column not in clean_values.columns]
    if missing:
        raise KeyError("Kunci duplikat tidak ditemukan setelah mapping: " + ", ".join(missing))

    rows: list[dict[str, Any]] = []
    group_counter = 1

    for column in duplicate_key_columns:
        work = clean_values[["ROW_ID", "SOURCE_ROW_NUMBER", column]].copy()
        normalized = work[column].map(_normalize_legacy_duplicate_value)
        eligible = normalized.ne("")
        repeated = eligible & normalized.duplicated(keep=False)
        if not repeated.any():
            continue

        repeated_values = normalized.loc[repeated]
        # sort=False mempertahankan urutan kemunculan dari file sumber.
        for normalized_value, indexes in repeated_values.groupby(repeated_values, sort=False).groups.items():
            index_list = list(indexes)
            group_id = f"DUP-{group_counter:05d}"
            group_counter += 1
            group_size = len(index_list)

            for order, index in enumerate(index_list):
                rows.append({
                    "ROW_ID": work.at[index, "ROW_ID"],
                    "SOURCE_ROW_NUMBER": work.at[index, "SOURCE_ROW_NUMBER"],
                    "COMPARISON_COLUMN": column,
                    "COMPARISON_VALUE": work.at[index, column],
                    "NORMALIZED_COMPARISON_VALUE": normalized_value,
                    "DUPLICATE_GROUP": group_id,
                    "GROUP_SIZE": group_size,
                    "STATUS": "ORIGINAL_GROUP_MEMBER" if order == 0 else "DUPLICATE",
                    "REASON": (
                        f"Nilai pada kolom '{column}' muncul {group_size} kali "
                        "setelah normalisasi pembanding"
                    ),
                })

    if not rows:
        return pd.DataFrame(columns=columns)

    return pd.DataFrame(rows, columns=columns).sort_values(
        ["COMPARISON_COLUMN", "DUPLICATE_GROUP", "SOURCE_ROW_NUMBER"],
        kind="stable",
    ).reset_index(drop=True)



def _safe_sort_values(
    df: pd.DataFrame,
    by: str | list[str],
) -> pd.DataFrame:
    """Mengurutkan kolom campuran tanpa mengubah nilai asli."""
    columns = [by] if isinstance(by, str) else list(by)
    existing = [column for column in columns if column in df.columns]
    if not existing or df.empty:
        return df.copy()

    work = df.copy()
    helper_columns: list[str] = []
    for position, column in enumerate(existing):
        helper_column = f"__SORT_KEY_{position}__"
        while helper_column in work.columns:
            helper_column += "_"
        work[helper_column] = (
            work[column].astype("string").fillna("").str.strip().str.casefold()
        )
        helper_columns.append(helper_column)
    work = work.sort_values(helper_columns, kind="stable", na_position="last")
    return work.drop(columns=helper_columns)


def _normalize_legacy_duplicate_value(value: Any) -> str:
    """Normalisasi pembanding duplikat yang sama dengan pipeline lama."""
    if is_missing_or_placeholder(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip().casefold())


def build_relationship_cardinality_report(df: pd.DataFrame) -> pd.DataFrame:
    """Mendeteksi satu nilai identitas yang terhubung ke banyak nilai lain.

    Ini berbeda dari duplikat baris. Contoh: satu CUSTOMER_ID memiliki beberapa
    Nomor Polis. Laporan dibuat untuk berbagai pasangan identitas yang tersedia
    pada dataset, bukan hanya CUSTOMER_ID dan Nomor Polis.
    """
    columns = [
        "ANCHOR_COLUMN", "ANCHOR_VALUE", "RELATED_COLUMN",
        "DISTINCT_RELATED_COUNT", "RELATED_VALUES", "ROW_IDS",
        "SOURCE_ROW_NUMBERS", "CONDITION", "REASON",
    ]

    candidate_pairs = [
        ("CUSTOMER_ID", "Nomor Polis"),
        ("CUSTOMER_ID", "POLICY_ID"),
        ("CUSTOMER_ID", "DEBITUR_ID"),
        ("CUSTOMER_ID", "NAMA_DEBITUR"),
        ("CUSTOMER_ID", "KTP_DEBITUR"),
        ("CUSTOMER_ID", "NPWP_DEBITUR"),
        ("DEBITUR_ID", "CUSTOMER_ID"),
        ("DEBITUR_ID", "Nomor Polis"),
        ("DEBITUR_ID", "POLICY_ID"),
        ("DEBITUR_ID", "KTP_DEBITUR"),
        ("DEBITUR_ID", "NPWP_DEBITUR"),
        ("Nomor Polis", "CUSTOMER_ID"),
        ("Nomor Polis", "POLICY_ID"),
        ("Nomor Polis", "Nama Pemegang Polis"),
        ("POLICY_ID", "CUSTOMER_ID"),
        ("POLICY_ID", "Nomor Polis"),
        ("KTP_DEBITUR", "DEBITUR_ID"),
        ("KTP_DEBITUR", "CUSTOMER_ID"),
        ("NPWP_DEBITUR", "DEBITUR_ID"),
        ("NPWP_DEBITUR", "CUSTOMER_ID"),
        ("Nomor Pelaporan", "Nama Pelanggan"),
        ("Nomor Pelaporan", "Nomor Polis"),
        ("Email pelanggan", "Nama Pelanggan"),
        ("Email pelanggan", "Nama Kontak"),
        ("WhatsApp", "Nama Pelanggan"),
        ("WhatsApp", "Nama Kontak"),
        ("Nomor Hp Pelanggan", "Nama Pelanggan"),
    ]

    rows: list[dict[str, Any]] = []
    for anchor_column, related_column in candidate_pairs:
        if anchor_column not in df.columns or related_column not in df.columns:
            continue

        base_columns = [anchor_column, related_column]
        for optional in ["ROW_ID", "SOURCE_ROW_NUMBER"]:
            if optional in df.columns:
                base_columns.append(optional)
        work = df[base_columns].copy()
        work["__ANCHOR_KEY__"] = work[anchor_column].map(_normalize_legacy_duplicate_value)
        work["__RELATED_KEY__"] = work[related_column].map(_normalize_legacy_duplicate_value)
        work = work[work["__ANCHOR_KEY__"].ne("") & work["__RELATED_KEY__"].ne("")]
        if work.empty:
            continue

        for _, group in work.groupby("__ANCHOR_KEY__", sort=False):
            distinct_related = group.drop_duplicates("__RELATED_KEY__", keep="first")
            count = len(distinct_related)
            if count <= 1:
                continue

            anchor_value = group.iloc[0][anchor_column]
            related_values = " | ".join(map(str, distinct_related[related_column].tolist()))
            row_ids = " | ".join(map(str, pd.unique(group["ROW_ID"].dropna()))) if "ROW_ID" in group else ""
            source_rows = " | ".join(map(str, pd.unique(group["SOURCE_ROW_NUMBER"].dropna()))) if "SOURCE_ROW_NUMBER" in group else ""
            rows.append({
                "ANCHOR_COLUMN": anchor_column,
                "ANCHOR_VALUE": anchor_value,
                "RELATED_COLUMN": related_column,
                "DISTINCT_RELATED_COUNT": count,
                "RELATED_VALUES": related_values,
                "ROW_IDS": row_ids,
                "SOURCE_ROW_NUMBERS": source_rows,
                "CONDITION": f"1 {anchor_column} memiliki >1 {related_column}",
                "REASON": (
                    f"Nilai '{anchor_value}' pada kolom '{anchor_column}' terhubung "
                    f"ke {count} nilai unik pada kolom '{related_column}'"
                ),
            })

    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows, columns=columns).sort_values(
        ["ANCHOR_COLUMN", "RELATED_COLUMN", "DISTINCT_RELATED_COUNT"],
        ascending=[True, True, False],
        kind="stable",
    ).reset_index(drop=True)


def build_legacy_duplicate_reports(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Mempertahankan seluruh aturan duplikat dari pipeline lama.

    Laporan ini berjalan berdampingan dengan ``build_duplicate_report`` yang
    memakai kunci pilihan pengguna. Dengan demikian, pengembangan konfigurasi
    manual tidak menghapus aturan bisnis lama untuk Nomor Pelaporan, responden,
    serta relasi CUSTOMER_ID-policy.
    """
    legacy_reports: dict[str, pd.DataFrame] = {}

    # 1. Nomor Pelaporan yang muncul lebih dari sekali.
    if "Nomor Pelaporan" in df.columns:
        values = df["Nomor Pelaporan"].map(_normalize_legacy_duplicate_value)
        duplicate_mask = values.ne("") & values.duplicated(keep=False)
        legacy_reports["Duplikat Nomor Pelaporan"] = _safe_sort_values(
            df.loc[duplicate_mask].copy(),
            "Nomor Pelaporan",
        )

        # Satu nomor pelaporan dipakai oleh lebih dari satu nama pelanggan.
        if "Nama Pelanggan" in df.columns:
            work = df.loc[values.ne(""), ["Nomor Pelaporan", "Nama Pelanggan"]].copy()
            work["__REPORT_KEY__"] = values.loc[values.ne("")]
            work["__NAME_KEY__"] = work["Nama Pelanggan"].map(_normalize_legacy_duplicate_value)
            conflict_keys = (
                work.loc[work["__NAME_KEY__"].ne("")]
                .groupby("__REPORT_KEY__", sort=False)["__NAME_KEY__"]
                .nunique(dropna=True)
                .loc[lambda series: series > 1]
                .index
            )
            conflicts = work[work["__REPORT_KEY__"].isin(conflict_keys)].copy()
            if conflicts.empty:
                legacy_reports["Konflik Nama per Pelaporan"] = pd.DataFrame(
                    columns=["Nomor Pelaporan", "DAFTAR_NAMA_PELANGGAN"]
                )
            else:
                legacy_reports["Konflik Nama per Pelaporan"] = (
                    conflicts.groupby("Nomor Pelaporan", sort=False)["Nama Pelanggan"]
                    .apply(lambda values: " | ".join(map(str, pd.unique(values.dropna()))))
                    .reset_index(name="DAFTAR_NAMA_PELANGGAN")
                )

    # 2. Duplikat responden berdasarkan aturan identitas lama.
    respondent_columns = [
        "Nomor Pelaporan", "Nomor Polis", "POLICY_ID", "Nama Pelanggan",
        "Nama Kontak", "Nama Perusahaan", "Email pelanggan",
        "Nomor Hp Pelanggan", "WhatsApp", "Kanwil", "Kantor Cabang",
        "Produk Yang Digunakan",
    ]
    available = [column for column in respondent_columns if column in df.columns]
    duplicate_mask = pd.Series(False, index=df.index, dtype=bool)
    duplicate_reasons = pd.Series("", index=df.index, dtype="string")

    if available:
        normalized = pd.DataFrame(
            {column: df[column].map(_normalize_legacy_duplicate_value) for column in available},
            index=df.index,
        )

        def append_reason(mask: pd.Series, reason: str) -> None:
            nonlocal duplicate_mask, duplicate_reasons
            mask = mask.reindex(df.index, fill_value=False)
            duplicate_mask |= mask
            current = duplicate_reasons.loc[mask].fillna("")
            duplicate_reasons.loc[mask] = current.apply(
                lambda old: reason if not old else f"{old}, {reason}"
            )

        for key_column in ["Nomor Pelaporan", "Nomor Polis", "POLICY_ID"]:
            if key_column in normalized.columns:
                values = normalized[key_column]
                append_reason(values.ne("") & values.duplicated(keep=False), key_column)

        for contact_column in ["Email pelanggan", "WhatsApp", "Nomor Hp Pelanggan"]:
            if contact_column in normalized.columns:
                values = normalized[contact_column]
                append_reason(values.ne("") & values.duplicated(keep=False), contact_column)

        if "Nama Perusahaan" in normalized.columns:
            for name_column in ["Nama Kontak", "Nama Pelanggan"]:
                if name_column not in normalized.columns:
                    continue
                pair = normalized[["Nama Perusahaan", name_column]]
                valid_pair = pair["Nama Perusahaan"].ne("") & pair[name_column].ne("")
                append_reason(
                    valid_pair & pair.duplicated(
                        subset=["Nama Perusahaan", name_column], keep=False
                    ),
                    f"Nama Perusahaan + {name_column}",
                )

        if {"Nomor Pelaporan", "Nomor Polis"}.issubset(normalized.columns):
            pair = normalized[["Nomor Pelaporan", "Nomor Polis"]]
            valid_pair = pair["Nomor Pelaporan"].ne("") & pair["Nomor Polis"].ne("")
            append_reason(
                valid_pair & pair.duplicated(
                    subset=["Nomor Pelaporan", "Nomor Polis"], keep=False
                ),
                "Nomor Pelaporan + Nomor Polis",
            )

    duplicate_respondents = df.loc[duplicate_mask].copy()
    if not duplicate_respondents.empty:
        duplicate_respondents.insert(
            0,
            "DUPLIKAT_BERDASARKAN",
            duplicate_reasons.loc[duplicate_mask].values,
        )
        sort_columns = [
            column for column in [
                "Nomor Pelaporan", "Nomor Polis", "POLICY_ID",
                "Nama Perusahaan", "Nama Kontak", "Nama Pelanggan",
                "Email pelanggan", "WhatsApp", "Nomor Hp Pelanggan",
                "Kantor Cabang",
            ] if column in duplicate_respondents.columns
        ]
        if sort_columns:
            duplicate_respondents = _safe_sort_values(duplicate_respondents, sort_columns)
    legacy_reports["Duplikat Responden Legacy"] = duplicate_respondents

    # 3. CUSTOMER_ID dengan lebih dari satu nomor polis, sama seperti pipeline lama.
    if "CUSTOMER_ID" in df.columns and "Nomor Polis" in df.columns:
        customer_values = df["CUSTOMER_ID"].map(_normalize_legacy_duplicate_value)
        policy_values = df["Nomor Polis"].map(_normalize_legacy_duplicate_value)
        eligible = customer_values.ne("") & policy_values.ne("")
        multi_policy_source = df.loc[eligible, ["CUSTOMER_ID", "Nomor Polis"]].copy()
        if multi_policy_source.empty:
            multi_policy = pd.DataFrame(columns=["CUSTOMER_ID", "JUMLAH_POLICY"])
        else:
            multi_policy = (
                multi_policy_source.groupby("CUSTOMER_ID", sort=False)["Nomor Polis"]
                .nunique(dropna=True)
                .reset_index(name="JUMLAH_POLICY")
            )
            multi_policy = multi_policy[multi_policy["JUMLAH_POLICY"] > 1]
        legacy_reports["Multi Policy"] = multi_policy

    if "CUSTOMER_ID" in df.columns and "POLICY_ID" in df.columns:
        customer_values = df["CUSTOMER_ID"].map(_normalize_legacy_duplicate_value)
        policy_values = df["POLICY_ID"].map(_normalize_legacy_duplicate_value)
        eligible = customer_values.ne("") & policy_values.ne("")
        detail_source = df.loc[eligible, ["CUSTOMER_ID", "POLICY_ID"]].copy()
        if detail_source.empty:
            detail = pd.DataFrame(columns=["CUSTOMER_ID", "POLICY_ID"])
        else:
            detail = (
                detail_source.groupby("CUSTOMER_ID", sort=False)["POLICY_ID"]
                .apply(lambda values: ", ".join(map(str, pd.unique(values.dropna()))))
                .reset_index(name="POLICY_ID")
            )
        legacy_reports["Multi Policy Detail"] = detail

        non_empty_policy = policy_values.ne("")
        duplicate_count = int(policy_values.loc[non_empty_policy].duplicated().sum())
        denominator = int(non_empty_policy.sum())
        duplicate_rate = duplicate_count / denominator * 100 if denominator else np.nan
        legacy_reports["Ringkasan Duplikat Policy"] = pd.DataFrame([{
            "Data duplicated": duplicate_count,
            "Duplicate Rate (%)": round(duplicate_rate, 2) if pd.notna(duplicate_rate) else np.nan,
        }])

    legacy_reports["Relasi Satu ke Banyak"] = build_relationship_cardinality_report(df)
    return legacy_reports


def build_quality_reports(
    df: pd.DataFrame,
    profile: str | None = None,
    transform_reports: dict[str, pd.DataFrame] | None = None,
    quality_columns: list[str] | None = None,
    header: int | None = 0,
    *,
    config: ETLConfig | None = None,
    cell_assessment: pd.DataFrame | None = None,
    duplicate_report: pd.DataFrame | None = None,
) -> dict[str, pd.DataFrame]:
    """Menyusun laporan kualitas baru sekaligus alias laporan lama."""
    if cell_assessment is None:
        if transform_reports and "Cell Assessment" in transform_reports:
            cell_assessment = transform_reports["Cell Assessment"]
        else:
            config = config or ETLConfig(input_file=".", output_file=".")
            work = add_stable_row_id(df, header)
            mapping = {
                column: recommend_column_type(column)
                for column in (quality_columns or list(df.columns))
                if column in work.columns and column not in {"ROW_ID", "SOURCE_ROW_NUMBER"}
            }
            _, generated = transform_data(work, config=config, type_mapping=mapping)
            cell_assessment = generated["Cell Assessment"]
    config = config or ETLConfig(input_file=".", output_file=".")
    duplicate_report = duplicate_report if duplicate_report is not None else pd.DataFrame()

    selected_columns = list(dict.fromkeys(cell_assessment["COLUMN"].tolist())) if not cell_assessment.empty else []
    total_rows = int(df["ROW_ID"].nunique()) if "ROW_ID" in df.columns else len(df)

    completeness_rows: list[dict[str, Any]] = []
    validity_rows: list[dict[str, Any]] = []
    for column in selected_columns:
        column_cells = cell_assessment[cell_assessment["COLUMN"] == column]
        empty_count = int(column_cells["ISSUE_CATEGORY"].eq("EMPTY").sum())
        placeholder_count = int(column_cells["ISSUE_CATEGORY"].eq("PLACEHOLDER").sum())
        filled_count = total_rows - empty_count - placeholder_count
        completeness_rows.append({
            "COLUMN": column,
            "TOTAL_ROWS": total_rows,
            "FILLED_COUNT": filled_count,
            "EMPTY_COUNT": empty_count,
            "PLACEHOLDER_COUNT": placeholder_count,
            "COMPLETENESS_PCT": round(filled_count / total_rows * 100, 2) if total_rows else np.nan,
        })

        assessed = column_cells["IS_VALID"].notna()
        valid_count = int(column_cells.loc[assessed, "IS_VALID"].eq(True).sum())
        invalid_count = int(column_cells.loc[assessed, "IS_VALID"].eq(False).sum())
        assessed_count = valid_count + invalid_count
        not_assessed = total_rows - assessed_count
        validity_rows.append({
            "COLUMN": column,
            "TOTAL_VALUES": total_rows,
            "ASSESSED_COUNT": assessed_count,
            "VALID_COUNT": valid_count,
            "INVALID_COUNT": invalid_count,
            "NOT_ASSESSED_COUNT": not_assessed,
            "VALIDITY_PCT": round(valid_count / assessed_count * 100, 2) if assessed_count else np.nan,
        })

    completeness = pd.DataFrame(completeness_rows, columns=[
        "COLUMN", "TOTAL_ROWS", "FILLED_COUNT", "EMPTY_COUNT",
        "PLACEHOLDER_COUNT", "COMPLETENESS_PCT",
    ])
    validity = pd.DataFrame(validity_rows, columns=[
        "COLUMN", "TOTAL_VALUES", "ASSESSED_COUNT", "VALID_COUNT",
        "INVALID_COUNT", "NOT_ASSESSED_COUNT", "VALIDITY_PCT",
    ])

    empty_data = cell_assessment[
        cell_assessment["ISSUE_CATEGORY"].isin(["EMPTY", "PLACEHOLDER"])
    ][[
        "ROW_ID", "SOURCE_ROW_NUMBER", "COLUMN", "ORIGINAL_VALUE",
        "ISSUE_CATEGORY", "VALIDATION_REASON",
    ]].rename(columns={
        "COLUMN": "COLUMN_NAME",
        "ISSUE_CATEGORY": "ISSUE_CATEGORY",
        "VALIDATION_REASON": "REASON",
    })

    invalid_data = cell_assessment[
        cell_assessment["ISSUE_CATEGORY"].isin(["INVALID", "NORMALIZATION_FAILED"])
    ][[
        "ROW_ID", "SOURCE_ROW_NUMBER", "COLUMN", "ORIGINAL_VALUE",
        "NORMALIZED_VALUE", "ISSUE_CATEGORY", "VALIDATION_REASON",
    ]].rename(columns={
        "COLUMN": "COLUMN_NAME",
        "VALIDATION_REASON": "INVALID_REASON",
    })

    placeholder_detail = empty_data[empty_data["ISSUE_CATEGORY"] == "PLACEHOLDER"].copy()
    audit = (
        transform_reports.get("Audit Perubahan", pd.DataFrame())
        if transform_reports else pd.DataFrame()
    )

    reports = {
        "Completeness Kolom": completeness,
        "Data Kosong": empty_data.reset_index(drop=True),
        "Validity Kolom": validity,
        "Data Tidak Valid": invalid_data.reset_index(drop=True),
        "Data Duplikat": duplicate_report.reset_index(drop=True),
        "Audit Perubahan": audit.reset_index(drop=True),
        # Alias lama agar integrasi lama tidak langsung rusak.
        "Completeness per Kolom": completeness.rename(columns={
            "COLUMN": "Kolom",
            "FILLED_COUNT": "Jumlah Terisi",
            "EMPTY_COUNT": "Jumlah Kosong",
            "COMPLETENESS_PCT": "Completeness (%)",
        }),
        "Detail Data Kosong": empty_data.reset_index(drop=True),
        "Detail Placeholder": placeholder_detail.reset_index(drop=True),
        "Validity per Kolom": validity.rename(columns={
            "COLUMN": "Kolom",
            "ASSESSED_COUNT": "Nilai Dinilai",
            "VALID_COUNT": "Jumlah Valid",
            "INVALID_COUNT": "Jumlah Tidak Valid",
            "VALIDITY_PCT": "Validity (%)",
        }),
        "Detail Data Tidak Valid": invalid_data.reset_index(drop=True),
        "Duplikat Responden": duplicate_report.reset_index(drop=True),
    }

    # Aturan duplikat lama tetap dijalankan berdampingan dengan kunci duplikat
    # pilihan pengguna. Nama ``Duplikat Responden`` dipertahankan untuk hasil
    # konfigurasi baru; hasil aturan lama disimpan pada sheet terpisah.
    legacy_duplicate_reports = build_legacy_duplicate_reports(df)
    reports.update(legacy_duplicate_reports)
    return reports


def apply_cleaning_rules(
    df: pd.DataFrame,
    config: ETLConfig,
    *,
    cell_assessment: pd.DataFrame | None = None,
    duplicate_report: pd.DataFrame | None = None,
    required_columns: list[str] | None = None,
    type_mapping: dict[str, str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Menerapkan cleaning konservatif dan membuat log terukur."""
    clean = df.copy()
    required_columns = required_columns or []
    type_mapping = type_mapping or {}
    cell_assessment = cell_assessment if cell_assessment is not None else pd.DataFrame()
    duplicate_report = duplicate_report if duplicate_report is not None else pd.DataFrame()

    reasons_by_row: dict[str, list[dict[str, Any]]] = {}

    def add_reason(row_id: str, reason: str, rule: str, trigger: str) -> None:
        reasons_by_row.setdefault(row_id, []).append({
            "reason": reason,
            "rule": rule,
            "trigger": trigger,
        })

    if config.remove_empty_required and required_columns and not cell_assessment.empty:
        required_issues = cell_assessment[
            cell_assessment["COLUMN"].isin(required_columns)
            & cell_assessment["ISSUE_CATEGORY"].isin(["EMPTY", "PLACEHOLDER"])
        ]
        for _, row in required_issues.iterrows():
            add_reason(
                str(row["ROW_ID"]),
                f"Kolom wajib '{row['COLUMN']}' kosong atau placeholder",
                "REMOVE_EMPTY_REQUIRED",
                f"{row['COLUMN']}={row['ORIGINAL_VALUE']}",
            )

    if not cell_assessment.empty:
        invalid = cell_assessment[cell_assessment["IS_VALID"].eq(False)]
        for _, row in invalid.iterrows():
            column_type = type_mapping.get(str(row["COLUMN"]), str(row["COLUMN_TYPE"]))
            should_remove = (
                (column_type == "phone" and config.remove_invalid_phone)
                or (column_type == "email" and config.remove_invalid_email)
                or (column_type not in {"phone", "email"} and config.remove_invalid_other)
            )
            if should_remove:
                add_reason(
                    str(row["ROW_ID"]),
                    f"Nilai invalid pada kolom '{row['COLUMN']}'",
                    f"REMOVE_INVALID_{column_type.upper()}",
                    f"{row['COLUMN']}={row['ORIGINAL_VALUE']} | {row['VALIDATION_REASON']}",
                )

    if config.remove_duplicates and not duplicate_report.empty:
        grouped = duplicate_report.groupby("DUPLICATE_GROUP", sort=False)
        for _, group in grouped:
            row_ids = group["ROW_ID"].tolist()
            if config.duplicate_keep == "first":
                removed_ids = row_ids[1:]
            elif config.duplicate_keep == "last":
                removed_ids = row_ids[:-1]
            else:
                removed_ids = row_ids
            for row_id in removed_ids:
                duplicate_row = group[group["ROW_ID"] == row_id].iloc[0]
                add_reason(
                    str(row_id),
                    "Baris termasuk kelompok duplikat",
                    "REMOVE_DUPLICATE",
                    f"{duplicate_row['COMPARISON_COLUMN']}={duplicate_row['COMPARISON_VALUE']}",
                )

    removal_rows: list[dict[str, Any]] = []
    for row_id, entries in reasons_by_row.items():
        source_row = clean.loc[clean["ROW_ID"] == row_id, "SOURCE_ROW_NUMBER"]
        removal_rows.append({
            "ROW_ID": row_id,
            "SOURCE_ROW_NUMBER": source_row.iloc[0] if not source_row.empty else pd.NA,
            "REMOVAL_REASON": " | ".join(dict.fromkeys(entry["reason"] for entry in entries)),
            "TRIGGER_RULE": " | ".join(dict.fromkeys(entry["rule"] for entry in entries)),
            "TRIGGER_VALUE": " | ".join(dict.fromkeys(entry["trigger"] for entry in entries)),
        })

    removal_log = pd.DataFrame(removal_rows, columns=[
        "ROW_ID", "SOURCE_ROW_NUMBER", "REMOVAL_REASON", "TRIGGER_RULE", "TRIGGER_VALUE",
    ])
    removed_ids = set(removal_log["ROW_ID"].tolist()) if not removal_log.empty else set()
    removed_data = clean[clean["ROW_ID"].isin(removed_ids)].copy()
    if not removed_data.empty:
        removed_data = removed_data.merge(removal_log, on=["ROW_ID", "SOURCE_ROW_NUMBER"], how="left")
    else:
        removed_data = pd.DataFrame(columns=[*clean.columns, "REMOVAL_REASON", "TRIGGER_RULE", "TRIGGER_VALUE"])
    clean = clean[~clean["ROW_ID"].isin(removed_ids)].reset_index(drop=True)
    return clean, removal_log, removed_data


# ============================================================
# KONFIGURASI, RINGKASAN, DAN EXPORT EXCEL
# ============================================================


def _serialize_config(config: ETLConfig, header_mapping: pd.DataFrame) -> pd.DataFrame:
    config_dict = asdict(config)
    config_dict.pop("progress_callback", None)
    config_dict["input_file"] = config.input_filename or Path(config.input_file).name
    config_dict["output_file"] = Path(config.output_file).name
    config_dict["pipeline_version"] = PIPELINE_VERSION
    rows: list[dict[str, Any]] = []
    for key, value in config_dict.items():
        if isinstance(value, (dict, list, tuple, set)):
            value = json.dumps(value, ensure_ascii=False, default=str)
        rows.append({"CONFIG_KEY": key, "CONFIG_VALUE": value})
    rows.append({
        "CONFIG_KEY": "resolved_header_mapping",
        "CONFIG_VALUE": header_mapping.to_json(orient="records", force_ascii=False),
    })
    return pd.DataFrame(rows)


def _build_process_summary(
    config: ETLConfig,
    started_at: datetime,
    finished_at: datetime,
    raw: pd.DataFrame,
    normalized: pd.DataFrame,
    clean: pd.DataFrame,
    removed: pd.DataFrame,
    reports: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    completeness = reports["Completeness Kolom"]
    validity = reports["Validity Kolom"]
    overall_completeness = (
        round(completeness["FILLED_COUNT"].sum() / completeness["TOTAL_ROWS"].sum() * 100, 2)
        if not completeness.empty and completeness["TOTAL_ROWS"].sum() else np.nan
    )
    overall_validity = (
        round(validity["VALID_COUNT"].sum() / validity["ASSESSED_COUNT"].sum() * 100, 2)
        if not validity.empty and validity["ASSESSED_COUNT"].sum() else np.nan
    )
    metrics = {
        "Input File": config.input_filename or Path(config.input_file).name,
        "Input Sheet": config.sheet_name,
        "Header Row": config.header,
        "Pipeline Mode": config.pipeline_mode,
        "Pipeline Version": PIPELINE_VERSION,
        "Started At": started_at.isoformat(timespec="seconds"),
        "Finished At": finished_at.isoformat(timespec="seconds"),
        "Execution Seconds": round((finished_at - started_at).total_seconds(), 3),
        "Initial Rows": len(raw),
        "Normalized Rows": len(normalized),
        "Clean Rows": len(clean),
        "Removed Rows": len(removed),
        # Satu baris dapat memiliki beberapa temuan invalid atau duplikat.
        # Metrik "Rows" selalu menghitung ROW_ID unik agar tidak pernah
        # melebihi jumlah baris sumber, sedangkan "Findings/Cells" menghitung
        # seluruh detail masalah untuk kebutuhan audit.
        "Duplicate Rows": (
            int(reports["Data Duplikat"]["ROW_ID"].nunique())
            if not reports["Data Duplikat"].empty and "ROW_ID" in reports["Data Duplikat"].columns
            else 0
        ),
        "Duplicate Findings": len(reports["Data Duplikat"]),
        "Empty Cells": int(reports["Data Kosong"]["ISSUE_CATEGORY"].eq("EMPTY").sum()) if not reports["Data Kosong"].empty else 0,
        "Placeholder Cells": int(reports["Data Kosong"]["ISSUE_CATEGORY"].eq("PLACEHOLDER").sum()) if not reports["Data Kosong"].empty else 0,
        "Invalid Rows": (
            int(reports["Data Tidak Valid"]["ROW_ID"].nunique())
            if not reports["Data Tidak Valid"].empty and "ROW_ID" in reports["Data Tidak Valid"].columns
            else 0
        ),
        "Invalid Cells": len(reports["Data Tidak Valid"]),
        "Audit Changes": len(reports["Audit Perubahan"]),
        "Overall Completeness (%)": overall_completeness,
        "Overall Validity (%)": overall_validity,
        "Process Status": "SUCCESS",
    }
    return pd.DataFrame({"METRIC": list(metrics.keys()), "VALUE": list(metrics.values())})


def _safe_sheet_name(name: str, used: set[str]) -> str:
    name = re.sub(r"[\[\]:*?/\\]", " ", str(name)).strip() or "Sheet"
    base = name[:31]
    candidate = base
    counter = 2
    while candidate in used:
        suffix = f" {counter}"
        candidate = base[: 31 - len(suffix)] + suffix
        counter += 1
    used.add(candidate)
    return candidate


def _excel_safe_value(value: Any) -> Any:
    if value is pd.NA or value is None:
        return None
    if isinstance(value, (dict, list, tuple, set)):
        return json.dumps(value, ensure_ascii=False, default=str)
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, pd.Timestamp):
        if value.tzinfo is not None:
            value = value.tz_localize(None)
        return value.to_pydatetime()
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and len(value) > 32767:
        return value[:32760] + "...[TRUNCATED]"
    return value


def _excel_safe_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    safe = df.copy()
    for column in safe.columns:
        safe[column] = safe[column].map(_excel_safe_value)
    return safe


def load_report(
    output_file: str | Path,
    normalized: pd.DataFrame,
    clean: pd.DataFrame,
    header_mapping: pd.DataFrame,
    removal_log: pd.DataFrame,
    reports: dict[str, pd.DataFrame],
    output_clean_sheet: str = "Data Bersih",
    *,
    raw: pd.DataFrame | None = None,
    removed: pd.DataFrame | None = None,
    configuration: pd.DataFrame | None = None,
    summary: pd.DataFrame | None = None,
    selected_output_sheets: list[str] | None = None,
) -> Path:
    output = Path(output_file)
    if output.suffix.lower() not in {".xlsx", ".xlsm"}:
        output = output.with_suffix(".xlsx")
    output.parent.mkdir(parents=True, exist_ok=True)

    raw = raw if raw is not None else pd.DataFrame()
    removed = removed if removed is not None else pd.DataFrame()
    configuration = configuration if configuration is not None else header_mapping
    summary = summary if summary is not None else pd.DataFrame(columns=["METRIC", "VALUE"])

    sheet_tables: dict[str, pd.DataFrame] = {
        "summary": summary,
        "config": configuration,
        "raw": raw,
        "normalized": normalized,
        "clean": clean,
        "removed": removed,
        "audit": reports.get("Audit Perubahan", pd.DataFrame(columns=[
            "ROW_ID", "SOURCE_ROW_NUMBER", "COLUMN", "ORIGINAL_VALUE",
            "NORMALIZED_VALUE", "CHANGE_TYPE", "VALIDATION_STATUS", "REASON",
        ])),
        "completeness": reports.get("Completeness Kolom", pd.DataFrame()),
        "missing": reports.get("Data Kosong", pd.DataFrame()),
        "validity": reports.get("Validity Kolom", pd.DataFrame()),
        "invalid": reports.get("Data Tidak Valid", pd.DataFrame()),
        "duplicates": reports.get("Data Duplikat", pd.DataFrame()),
        "duplicate_report_number": reports.get("Duplikat Nomor Pelaporan", pd.DataFrame()),
        "report_name_conflict": reports.get("Konflik Nama per Pelaporan", pd.DataFrame()),
        "legacy_duplicates": reports.get("Duplikat Responden Legacy", pd.DataFrame()),
        "multi_policy": reports.get("Multi Policy", pd.DataFrame()),
        "multi_policy_detail": reports.get("Multi Policy Detail", pd.DataFrame()),
        "duplicate_policy_summary": reports.get("Ringkasan Duplikat Policy", pd.DataFrame()),
        "relations": reports.get("Relasi Satu ke Banyak", pd.DataFrame()),
        "header_mapping": header_mapping,
        "removal_log": removal_log,
    }
    selected_keys = list(dict.fromkeys(selected_output_sheets or DEFAULT_OUTPUT_SHEETS))
    unknown_keys = [key for key in selected_keys if key not in REPORT_REGISTRY]
    if unknown_keys:
        raise ValueError("Sheet output tidak dikenali: " + ", ".join(unknown_keys))
    if not selected_keys:
        raise ValueError("Minimal satu sheet output harus dipilih.")
    expected_sheets = [
        (output_clean_sheet if key == "clean" else REPORT_REGISTRY[key], sheet_tables[key])
        for key in selected_keys
    ]

    used: set[str] = set()
    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for requested_name, table in expected_sheets:
                safe_table = _excel_safe_dataframe(table)
                safe_table.to_excel(
                    writer,
                    sheet_name=_safe_sheet_name(requested_name, used),
                    index=False,
                )

            workbook = writer.book
            for worksheet in workbook.worksheets:
                worksheet.freeze_panes = "A2"
                if worksheet.max_row >= 1 and worksheet.max_column >= 1:
                    worksheet.auto_filter.ref = worksheet.dimensions
                    for cell in worksheet[1]:
                        cell.font = Font(bold=True)
                for column_cells in worksheet.iter_cols():
                    sample = list(column_cells[: min(len(column_cells), 300)])
                    values = [str(cell.value) if cell.value is not None else "" for cell in sample]
                    width = min(max(max((len(value) for value in values), default=0) + 2, 10), 45)
                    worksheet.column_dimensions[column_cells[0].column_letter].width = width

                headers = {cell.value: cell.column for cell in worksheet[1]}
                for header_name, column_index in headers.items():
                    if isinstance(header_name, str) and (
                        header_name.endswith("_PCT") or header_name.endswith("(%)")
                    ):
                        for row_index in range(2, worksheet.max_row + 1):
                            worksheet.cell(row=row_index, column=column_index).number_format = "0.00"
    except PermissionError as exc:
        raise PermissionError(
            f"File output tidak dapat ditulis. Tutup file Excel jika sedang dibuka: {output}"
        ) from exc
    except MemoryError as exc:
        raise MemoryError("Memori tidak cukup untuk menulis laporan Excel.") from exc
    except Exception as exc:
        raise RuntimeError(f"Kegagalan penulisan Excel: {exc}") from exc
    return output


# ============================================================
# ORKESTRASI PIPELINE
# ============================================================


def run_etl(config: ETLConfig) -> dict[str, Any]:
    def progress(percent: int, message: str) -> None:
        if config.progress_callback is not None:
            config.progress_callback(min(max(percent, 0), 100), message)

    started_at = datetime.now()
    progress(0, "Memvalidasi konfigurasi")

    raw_source = extract_data(
        file_path=config.input_file,
        sheet_name=config.sheet_name,
        header=config.header,
        selected_columns=config.selected_columns,
    )
    errors, warnings = validate_etl_config(config, raw_source.columns)
    if errors:
        raise ValueError("Konfigurasi tidak valid: " + " | ".join(errors))
    progress(12, "Data sumber berhasil dibaca")

    raw = add_stable_row_id(raw_source, config.header)
    harmonized, header_mapping = harmonize_headers(
        raw,
        {**config.canonical_mapping, "ROW_ID": "ROW_ID", "SOURCE_ROW_NUMBER": "SOURCE_ROW_NUMBER"},
    )
    type_mapping, required_columns, duplicate_keys = _effective_config_mappings(config, header_mapping)
    progress(25, "Mapping kolom diterapkan")

    normalized, transform_reports = transform_data(
        harmonized,
        config=config,
        type_mapping=type_mapping,
    )
    clean_values = transform_reports["Clean Values"]
    cell_assessment = transform_reports["Cell Assessment"]
    progress(52, "Normalisasi dan validasi selesai")

    duplicate_report = build_duplicate_report(clean_values, duplicate_keys)
    reports = build_quality_reports(
        harmonized,
        config=config,
        transform_reports=transform_reports,
        cell_assessment=cell_assessment,
        duplicate_report=duplicate_report,
    )
    progress(70, "Laporan kualitas disusun")

    clean, removal_log, removed = apply_cleaning_rules(
        clean_values,
        config,
        cell_assessment=cell_assessment,
        duplicate_report=duplicate_report,
        required_columns=required_columns,
        type_mapping=type_mapping,
    )
    progress(84, "Aturan cleaning diterapkan")

    configuration = _serialize_config(config, header_mapping)
    finished_at = datetime.now()
    summary = _build_process_summary(
        config,
        started_at,
        finished_at,
        raw,
        normalized,
        clean,
        removed,
        reports,
    )

    if not config.create_detail_reports:
        reports = dict(reports)
        detail_keys = [
            "Data Kosong", "Data Tidak Valid", "Data Duplikat",
            "Detail Data Kosong", "Detail Placeholder",
            "Detail Data Tidak Valid", "Duplikat Responden",
        ]
        for key in detail_keys:
            if key in reports:
                reports[key] = reports[key].iloc[0:0].copy()

    reports["Ringkasan"] = pd.DataFrame([{
        "Jumlah Baris": len(raw),
        "Jumlah Kolom": len(config.selected_columns),
        "Completeness (%)": _metric_from_summary(summary, "Overall Completeness (%)"),
        "Validity (%)": _metric_from_summary(summary, "Overall Validity (%)"),
        "Konfigurasi Pipeline": config.pipeline_mode,
    }])

    output = load_report(
        config.output_file,
        normalized,
        clean,
        header_mapping,
        removal_log,
        reports,
        output_clean_sheet=config.output_clean_sheet,
        raw=raw,
        removed=removed,
        configuration=configuration,
        summary=summary,
        selected_output_sheets=config.output_sheets,
    )
    progress(100, "Proses ETL selesai")

    return {
        "raw": raw,
        "harmonized": harmonized,
        "normalized": normalized,
        "clean": clean,
        "removed": removed,
        "header_mapping": header_mapping,
        "removal_log": removal_log,
        "configuration": configuration,
        "summary": summary,
        "reports": reports,
        "warnings": warnings,
        "pipeline_mode": config.pipeline_mode,
        "profile": config.pipeline_mode,  # kompatibilitas key lama, tanpa deteksi otomatis
        "output_file": output,
    }


def _metric_from_summary(summary: pd.DataFrame, metric: str) -> Any:
    match = summary.loc[summary["METRIC"] == metric, "VALUE"]
    return match.iloc[0] if not match.empty else np.nan


def sanitize_filename(filename: str, default: str = "hasil_etl.xlsx") -> str:
    name = Path(filename).name
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._")
    return name or default


# ============================================================
# CLI
# ============================================================


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Pipeline ETL dan data-quality Askrindo.")
    parser.add_argument("input_file")
    parser.add_argument("output_file")
    parser.add_argument("--sheet", default=0)
    parser.add_argument("--header", type=int, default=0)
    parser.add_argument("--mode", choices=["Generic"], default="Generic")
    parser.add_argument("--columns", nargs="+", required=True)
    parser.add_argument("--mapping-json", required=True, help='JSON tipe kolom, contoh: {"Email":"email"}')
    parser.add_argument("--canonical-json", default="{}")
    parser.add_argument("--required", nargs="*", default=[])
    parser.add_argument("--duplicate-keys", nargs="*", default=[])
    parser.add_argument("--remove-empty-required", action="store_true")
    parser.add_argument("--remove-invalid-phone", action="store_true")
    parser.add_argument("--remove-invalid-email", action="store_true")
    parser.add_argument("--remove-invalid-other", action="store_true")
    parser.add_argument("--remove-duplicates", action="store_true")
    parser.add_argument("--output-sheets", nargs="*", default=list(DEFAULT_OUTPUT_SHEETS), choices=list(REPORT_REGISTRY))
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    sheet: str | int = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
    result = run_etl(ETLConfig(
        input_file=args.input_file,
        output_file=args.output_file,
        input_filename=Path(args.input_file).name,
        sheet_name=sheet,
        header=args.header,
        selected_columns=args.columns,
        pipeline_mode=args.mode,
        column_mapping=json.loads(args.mapping_json),
        canonical_mapping=json.loads(args.canonical_json),
        required_columns=args.required,
        duplicate_key_columns=args.duplicate_keys,
        remove_empty_required=args.remove_empty_required,
        remove_invalid_phone=args.remove_invalid_phone,
        remove_invalid_email=args.remove_invalid_email,
        remove_invalid_other=args.remove_invalid_other,
        remove_duplicates=args.remove_duplicates,
        output_sheets=args.output_sheets,
    ))
    print(f"Mode pipeline : {result['pipeline_mode']}")
    print(f"Baris awal    : {len(result['raw'])}")
    print(f"Baris bersih  : {len(result['clean'])}")
    print(f"Output        : {result['output_file']}")
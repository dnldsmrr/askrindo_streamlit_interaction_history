from __future__ import annotations

import hashlib
import tempfile
import traceback
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from interaction_pipeline import (
    DEFAULT_OUTPUT_SHEETS,
    INTERACTION_METHODS,
    INTERACTION_PIPELINE_VERSION,
    InteractionConfig,
    export_interaction_excel,
    get_sheet_names,
    read_excel_bytes,
    run_interaction_pipeline,
    safe_output_filename,
    validate_interaction_config,
)

st.set_page_config(page_title="Riwayat Interaksi Pelanggan", layout="wide")
st.title("Riwayat Interaksi Pelanggan")
st.caption(
    "Mencocokkan data Contact Centre ke ROW_ID master pelanggan, lalu menghitung "
    "berapa kali dan kapan setiap pelanggan menghubungi perusahaan."
)

PREFIX = "interaction_"
OPTIONAL = "Tidak digunakan"


def state_key(name: str) -> str:
    return f"{PREFIX}{name}"


def reset_interaction_result() -> None:
    for key in [
        state_key("result"),
        state_key("excel"),
        state_key("filename"),
        state_key("debug_error"),
    ]:
        st.session_state.pop(key, None)


@st.cache_data(show_spinner=False)
def cached_sheet_names(file_bytes: bytes) -> list[str]:
    return get_sheet_names(file_bytes)


@st.cache_data(show_spinner=False)
def cached_read_excel(
    file_bytes: bytes,
    sheet_name: str,
    header_row: int,
) -> pd.DataFrame:
    return read_excel_bytes(file_bytes, sheet_name, header_row)


def safe_dataframe(df: pd.DataFrame, **kwargs: Any) -> None:
    try:
        st.dataframe(df, **kwargs)
    except (TypeError, ValueError, OverflowError):
        st.dataframe(df.astype("string"), **kwargs)


def optional_selectbox(
    label: str,
    columns: list[str],
    *,
    key: str,
    help_text: str | None = None,
) -> str | None:
    selected = st.selectbox(
        label,
        [OPTIONAL, *columns],
        key=key,
        help=help_text,
        on_change=reset_interaction_result,
    )
    return None if selected == OPTIONAL else selected


def metric_value(summary: pd.DataFrame, name: str, default: Any = 0) -> Any:
    if summary.empty or not {"METRIC", "VALUE"}.issubset(summary.columns):
        return default
    values = summary.loc[summary["METRIC"] == name, "VALUE"]
    return values.iloc[0] if not values.empty else default


def file_hash(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


# ============================================================
# A. MASTER PELANGGAN
# ============================================================
st.header("A. Upload Master Pelanggan")
master_file = st.file_uploader(
    "File master pelanggan",
    type=["xlsx", "xlsm", "xls"],
    key=state_key("master_file"),
    on_change=reset_interaction_result,
    help="Master harus memiliki ROW_ID dan minimal satu sumber kontak: telepon atau email.",
)
if master_file is None:
    st.info("Unggah master pelanggan untuk memulai.")
    st.stop()

master_bytes = master_file.getvalue()
if not master_bytes:
    st.error("File master pelanggan kosong.")
    st.stop()

try:
    master_sheets = cached_sheet_names(master_bytes)
except Exception as exc:
    st.error(f"Master pelanggan tidak dapat dibaca. Detail: {exc}")
    st.stop()

master_structure = st.columns([2, 1])
master_sheet = master_structure[0].selectbox(
    "Sheet master",
    master_sheets,
    key=state_key("master_sheet"),
    on_change=reset_interaction_result,
)
master_header = int(
    master_structure[1].number_input(
        "Header master, dimulai dari 0",
        min_value=0,
        value=0,
        step=1,
        key=state_key("master_header"),
        on_change=reset_interaction_result,
    )
)

try:
    master_df = cached_read_excel(master_bytes, master_sheet, master_header)
except Exception as exc:
    st.error(f"Master gagal dibaca dengan sheet/header tersebut. Detail: {exc}")
    st.stop()

if master_df.empty or len(master_df.columns) == 0:
    st.error("Master tidak memiliki data atau header yang valid.")
    st.stop()
master_columns = [str(column) for column in master_df.columns]

master_map_1, master_map_2 = st.columns(2)
with master_map_1:
    master_row_id_column = st.selectbox(
        "Kolom ROW_ID",
        master_columns,
        key=state_key("master_row_id"),
        on_change=reset_interaction_result,
    )
    master_customer_name_column = st.selectbox(
        "Kolom nama pelanggan",
        master_columns,
        key=state_key("master_customer_name"),
        on_change=reset_interaction_result,
    )
    master_company_name_column = st.selectbox(
        "Kolom nama perusahaan",
        master_columns,
        key=state_key("master_company_name"),
        on_change=reset_interaction_result,
    )
with master_map_2:
    master_phone_columns = st.multiselect(
        "Kolom telepon master",
        master_columns,
        key=state_key("master_phone_columns"),
        on_change=reset_interaction_result,
        help="Dapat memilih WhatsApp, nomor handphone, telepon kantor, dan kolom telepon lain.",
    )
    master_email_columns = st.multiselect(
        "Kolom email master",
        master_columns,
        key=state_key("master_email_columns"),
        on_change=reset_interaction_result,
    )

master_metrics = st.columns(3)
master_metrics[0].metric("Nama file", master_file.name)
master_metrics[1].metric("Jumlah baris", len(master_df))
master_metrics[2].metric("Jumlah kolom", len(master_df.columns))
with st.expander("Preview master pelanggan", expanded=False):
    safe_dataframe(master_df.head(10), use_container_width=True, hide_index=True)


# ============================================================
# B. CONTACT CENTRE
# ============================================================
st.header("B. Upload Data Contact Centre")
contact_file = st.file_uploader(
    "File data Contact Centre",
    type=["xlsx", "xlsm", "xls"],
    key=state_key("contact_file"),
    on_change=reset_interaction_result,
)
if contact_file is None:
    st.info("Unggah data Contact Centre untuk melanjutkan.")
    st.stop()

contact_bytes = contact_file.getvalue()
if not contact_bytes:
    st.error("File Contact Centre kosong.")
    st.stop()

try:
    contact_sheets = cached_sheet_names(contact_bytes)
except Exception as exc:
    st.error(f"File Contact Centre tidak dapat dibaca. Detail: {exc}")
    st.stop()

contact_structure = st.columns([2, 1])
contact_sheet = contact_structure[0].selectbox(
    "Sheet Contact Centre",
    contact_sheets,
    key=state_key("contact_sheet"),
    on_change=reset_interaction_result,
)
contact_header = int(
    contact_structure[1].number_input(
        "Header Contact Centre, dimulai dari 0",
        min_value=0,
        value=0,
        step=1,
        key=state_key("contact_header"),
        on_change=reset_interaction_result,
    )
)

try:
    contact_df = cached_read_excel(contact_bytes, contact_sheet, contact_header)
except Exception as exc:
    st.error(f"Contact Centre gagal dibaca dengan sheet/header tersebut. Detail: {exc}")
    st.stop()

if contact_df.empty or len(contact_df.columns) == 0:
    st.error("Data Contact Centre tidak memiliki data atau header yang valid.")
    st.stop()
contact_columns = [str(column) for column in contact_df.columns]

contact_map_1, contact_map_2, contact_map_3 = st.columns(3)
with contact_map_1:
    contact_date_column = optional_selectbox(
        "Kolom tanggal kontak",
        contact_columns,
        key=state_key("contact_date"),
    )
    contact_report_number_column = optional_selectbox(
        "Kolom nomor pelaporan",
        contact_columns,
        key=state_key("contact_report_number"),
    )
    contact_customer_name_column = optional_selectbox(
        "Kolom nama pelanggan",
        contact_columns,
        key=state_key("contact_customer_name"),
    )
    contact_company_name_column = optional_selectbox(
        "Kolom nama perusahaan",
        contact_columns,
        key=state_key("contact_company_name"),
    )
with contact_map_2:
    contact_phone_columns = st.multiselect(
        "Kolom telepon Contact Centre",
        contact_columns,
        key=state_key("contact_phone_columns"),
        on_change=reset_interaction_result,
    )
    contact_email_columns = st.multiselect(
        "Kolom email Contact Centre",
        contact_columns,
        key=state_key("contact_email_columns"),
        on_change=reset_interaction_result,
    )
    contact_channel_column = optional_selectbox(
        "Kolom saluran",
        contact_columns,
        key=state_key("contact_channel"),
    )
    contact_category_column = optional_selectbox(
        "Kolom kategori",
        contact_columns,
        key=state_key("contact_category"),
    )
with contact_map_3:
    contact_service_type_column = optional_selectbox(
        "Kolom jenis layanan",
        contact_columns,
        key=state_key("contact_service"),
    )
    contact_status_column = optional_selectbox(
        "Kolom status",
        contact_columns,
        key=state_key("contact_status"),
    )
    contact_description_column = optional_selectbox(
        "Kolom keterangan",
        contact_columns,
        key=state_key("contact_description"),
    )

contact_metrics = st.columns(3)
contact_metrics[0].metric("Nama file", contact_file.name)
contact_metrics[1].metric("Jumlah baris", len(contact_df))
contact_metrics[2].metric("Jumlah kolom", len(contact_df.columns))
with st.expander("Preview data Contact Centre", expanded=True):
    safe_dataframe(contact_df.head(10), use_container_width=True, hide_index=True)


# ============================================================
# C. ATURAN MATCHING
# ============================================================
st.header("C. Aturan Matching")
match_cols = st.columns(4)
match_phone = match_cols[0].checkbox(
    "Gunakan telepon",
    value=True,
    key=state_key("match_phone"),
    on_change=reset_interaction_result,
)
match_email = match_cols[1].checkbox(
    "Gunakan email",
    value=True,
    key=state_key("match_email"),
    on_change=reset_interaction_result,
)
match_name = match_cols[2].checkbox(
    "Gunakan nama pelanggan",
    value=True,
    key=state_key("match_name"),
    on_change=reset_interaction_result,
)
match_company = match_cols[3].checkbox(
    "Gunakan nama perusahaan",
    value=True,
    key=state_key("match_company"),
    on_change=reset_interaction_result,
)
allow_name_company_fallback = st.checkbox(
    "Izinkan fallback exact match nama pelanggan + nama perusahaan",
    value=True,
    key=state_key("allow_name_company_fallback"),
    on_change=reset_interaction_result,
    help=(
        "Fallback hanya digunakan ketika telepon dan email tidak menemukan kandidat. "
        "Nama atau perusahaan tidak pernah digunakan sendiri."
    ),
)
st.info(
    "Prioritas: telepon+email, telepon, email, lalu nama+perusahaan sebagai fallback atau pendukung. "
    "Sistem tidak memilih kandidat pertama jika hasil masih ambigu."
)


# ============================================================
# D. DEFINISI INTERAKSI
# ============================================================
st.header("D. Definisi Satu Interaksi")
interaction_label_to_key = {label: key for key, label in INTERACTION_METHODS.items()}
selected_interaction_label = st.radio(
    "Dasar perhitungan jumlah interaksi",
    options=list(interaction_label_to_key.keys()),
    index=0,
    key=state_key("interaction_definition_label"),
    on_change=reset_interaction_result,
)
interaction_definition = interaction_label_to_key[selected_interaction_label]

if interaction_definition == "report_number" and contact_report_number_column is None:
    st.warning(
        "Mode nomor pelaporan unik dipilih, tetapi kolom nomor pelaporan belum ditentukan. "
        "Baris tanpa nomor pelaporan akan dihitung berdasarkan CONTACT_ROW_ID."
    )
if interaction_definition == "date_channel_category":
    missing_combination = [
        label
        for label, column in [
            ("tanggal", contact_date_column),
            ("saluran", contact_channel_column),
            ("kategori", contact_category_column),
        ]
        if column is None
    ]
    if missing_combination:
        st.warning(
            "Mode kombinasi membutuhkan mapping: " + ", ".join(missing_combination)
            + ". Nilai kosong tetap menjadi bagian dari kunci sehingga hasil harus ditinjau."
        )


# ============================================================
# E. OUTPUT
# ============================================================
st.header("E. Pilih Sheet Output")
selected_output_sheets = st.multiselect(
    "Sheet yang ditulis ke Excel",
    options=DEFAULT_OUTPUT_SHEETS,
    default=[
        "Ringkasan Proses",
        "Konfigurasi",
        "Hasil Matching",
        "Ringkasan per Pelanggan",
        "Detail Riwayat Kontak",
        "Data Ambigu",
        "Data Tidak Ditemukan",
    ],
    key=state_key("output_sheets"),
    on_change=reset_interaction_result,
    help="Analisis tetap dijalankan penuh. Pilihan ini hanya mengatur sheet yang diekspor.",
)


# ============================================================
# F. VALIDASI DAN PREVIEW
# ============================================================
config = InteractionConfig(
    master_sheet=master_sheet,
    master_header=master_header,
    master_row_id_column=master_row_id_column,
    master_customer_name_column=master_customer_name_column,
    master_company_name_column=master_company_name_column,
    master_phone_columns=master_phone_columns,
    master_email_columns=master_email_columns,
    contact_sheet=contact_sheet,
    contact_header=contact_header,
    contact_date_column=contact_date_column,
    contact_report_number_column=contact_report_number_column,
    contact_customer_name_column=contact_customer_name_column,
    contact_company_name_column=contact_company_name_column,
    contact_phone_columns=contact_phone_columns,
    contact_email_columns=contact_email_columns,
    contact_channel_column=contact_channel_column,
    contact_category_column=contact_category_column,
    contact_service_type_column=contact_service_type_column,
    contact_status_column=contact_status_column,
    contact_description_column=contact_description_column,
    match_phone=match_phone,
    match_email=match_email,
    match_name=match_name,
    match_company=match_company,
    allow_name_company_fallback=allow_name_company_fallback,
    interaction_definition=interaction_definition,
    output_sheets=selected_output_sheets,
    master_filename=master_file.name,
    contact_filename=contact_file.name,
)

validation_errors = validate_interaction_config(config, master_columns, contact_columns)
if not selected_output_sheets:
    validation_errors.append("Minimal satu sheet output harus dipilih.")
if master_row_id_column in {master_customer_name_column, master_company_name_column}:
    validation_errors.append("Kolom ROW_ID tidak boleh sama dengan kolom nama pelanggan atau perusahaan.")
if not match_phone and not match_email and not allow_name_company_fallback:
    validation_errors.append("Aktifkan minimal satu jalur matching.")
if match_phone and (not master_phone_columns or not contact_phone_columns):
    validation_errors.append("Matching telepon aktif, tetapi mapping telepon master atau Contact Centre belum lengkap.")
if match_email and (not master_email_columns or not contact_email_columns):
    validation_errors.append("Matching email aktif, tetapi mapping email master atau Contact Centre belum lengkap.")
if allow_name_company_fallback and (
    contact_customer_name_column is None or contact_company_name_column is None
):
    validation_errors.append("Fallback nama+perusahaan membutuhkan mapping kedua kolom pada Contact Centre.")

st.header("F. Preview Konfigurasi")
config_preview = pd.DataFrame([
    {"BAGIAN": "Master", "PENGATURAN": "ROW_ID", "NILAI": master_row_id_column},
    {"BAGIAN": "Master", "PENGATURAN": "Nama pelanggan", "NILAI": master_customer_name_column},
    {"BAGIAN": "Master", "PENGATURAN": "Nama perusahaan", "NILAI": master_company_name_column},
    {"BAGIAN": "Master", "PENGATURAN": "Telepon", "NILAI": "; ".join(master_phone_columns)},
    {"BAGIAN": "Master", "PENGATURAN": "Email", "NILAI": "; ".join(master_email_columns)},
    {"BAGIAN": "Contact Centre", "PENGATURAN": "Tanggal", "NILAI": contact_date_column or OPTIONAL},
    {"BAGIAN": "Contact Centre", "PENGATURAN": "Nomor pelaporan", "NILAI": contact_report_number_column or OPTIONAL},
    {"BAGIAN": "Contact Centre", "PENGATURAN": "Nama pelanggan", "NILAI": contact_customer_name_column or OPTIONAL},
    {"BAGIAN": "Contact Centre", "PENGATURAN": "Nama perusahaan", "NILAI": contact_company_name_column or OPTIONAL},
    {"BAGIAN": "Contact Centre", "PENGATURAN": "Telepon", "NILAI": "; ".join(contact_phone_columns)},
    {"BAGIAN": "Contact Centre", "PENGATURAN": "Email", "NILAI": "; ".join(contact_email_columns)},
    {"BAGIAN": "Matching", "PENGATURAN": "Fallback nama+perusahaan", "NILAI": allow_name_company_fallback},
    {"BAGIAN": "Interaksi", "PENGATURAN": "Definisi", "NILAI": selected_interaction_label},
    {"BAGIAN": "Sistem", "PENGATURAN": "Versi pipeline", "NILAI": INTERACTION_PIPELINE_VERSION},
])
safe_dataframe(config_preview, use_container_width=True, hide_index=True)
for error in validation_errors:
    st.error(error)


# ============================================================
# G. PROSES
# ============================================================
st.header("G. Proses")
run_button = st.button(
    "Jalankan Analisis Riwayat Pelanggan",
    type="primary",
    use_container_width=True,
    disabled=bool(validation_errors),
)

if run_button:
    progress = st.progress(0)
    status = st.empty()
    try:
        status.write("Memvalidasi dan menormalisasi master pelanggan")
        progress.progress(15)
        result = run_interaction_pipeline(master_df, contact_df, config)
        status.write("Menyusun riwayat dan ringkasan per pelanggan")
        progress.progress(75)

        output_filename = safe_output_filename(contact_file.name)
        with tempfile.TemporaryDirectory(prefix="askrindo_interaction_") as temp_dir:
            output_path = Path(temp_dir) / output_filename
            export_interaction_excel(result, output_path)
            if not output_path.exists() or output_path.stat().st_size == 0:
                raise RuntimeError("Pipeline tidak menghasilkan file Excel.")
            st.session_state[state_key("excel")] = output_path.read_bytes()

        st.session_state[state_key("filename")] = output_filename
        st.session_state[state_key("result")] = result
        st.session_state.pop(state_key("debug_error"), None)
        progress.progress(100)
        status.write("Analisis riwayat pelanggan selesai")
        st.success("Analisis selesai.")
    except MemoryError:
        st.session_state[state_key("debug_error")] = traceback.format_exc()
        st.error("Memori tidak cukup. Kurangi ukuran file atau jumlah kolom.")
    except PermissionError:
        st.session_state[state_key("debug_error")] = traceback.format_exc()
        st.error("File output tidak dapat ditulis. Pastikan file Excel hasil tidak sedang dibuka.")
    except Exception as exc:
        st.session_state[state_key("debug_error")] = traceback.format_exc()
        st.error(
            "Proses gagal. Periksa keunikan ROW_ID, mapping kolom, dan format file. "
            f"Detail: {exc}"
        )


# ============================================================
# H. HASIL
# ============================================================
if state_key("result") in st.session_state and state_key("excel") in st.session_state:
    st.header("H. Hasil")
    result = st.session_state[state_key("result")]
    summary = result["summary"]

    metrics_row_1 = st.columns(5)
    metrics_row_1[0].metric("Master pelanggan", metric_value(summary, "Jumlah master pelanggan"))
    metrics_row_1[1].metric("Baris kontak", metric_value(summary, "Jumlah baris Contact Centre"))
    metrics_row_1[2].metric("Matched", metric_value(summary, "Jumlah MATCHED"))
    metrics_row_1[3].metric("Ambiguous", metric_value(summary, "Jumlah AMBIGUOUS"))
    metrics_row_1[4].metric("Not found", metric_value(summary, "Jumlah NOT_FOUND"))

    metrics_row_2 = st.columns(4)
    metrics_row_2[0].metric(
        "Pelanggan pernah menghubungi",
        metric_value(summary, "Jumlah pelanggan pernah menghubungi"),
    )
    metrics_row_2[1].metric(
        "Pelanggan belum menghubungi",
        metric_value(summary, "Jumlah pelanggan belum menghubungi"),
    )
    metrics_row_2[2].metric("Total interaksi", metric_value(summary, "Total interaksi"))
    completeness = metric_value(summary, "Kompleteness matching (%)", 0)
    metrics_row_2[3].metric("Completeness matching", f"{float(completeness):.2f}%")

    tabs = st.tabs([
        "Ringkasan Pelanggan",
        "Detail Riwayat",
        "Hasil Matching",
        "Ambiguous",
        "Not Found",
        "Ringkasan Proses",
    ])
    with tabs[0]:
        safe_dataframe(result["customer_summary"].head(1000), use_container_width=True, hide_index=True)
    with tabs[1]:
        safe_dataframe(result["detail"].head(1000), use_container_width=True, hide_index=True)
    with tabs[2]:
        safe_dataframe(result["result"].head(1000), use_container_width=True, hide_index=True)
    with tabs[3]:
        safe_dataframe(result["ambiguous"].head(1000), use_container_width=True, hide_index=True)
    with tabs[4]:
        safe_dataframe(result["not_found"].head(1000), use_container_width=True, hide_index=True)
    with tabs[5]:
        safe_dataframe(summary, use_container_width=True, hide_index=True)

    st.download_button(
        "Unduh hasil Excel",
        data=st.session_state[state_key("excel")],
        file_name=st.session_state[state_key("filename")],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

from __future__ import annotations

import hashlib
import io
import json
import tempfile
import traceback
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from pipeline_etl_askrindo import (
    COLUMN_TYPE_LABELS,
    DEFAULT_OUTPUT_SHEETS,
    DEFAULT_PLACEHOLDERS,
    ETLConfig,
    PIPELINE_VERSION,
    REPORT_REGISTRY,
    build_mapping_recommendations,
    get_excel_sheet_names,
    read_preview_data,
    run_etl,
    sanitize_filename,
    validate_etl_config,
)

st.set_page_config(page_title="Data Quality Testing", layout="wide")
st.title("Data Quality Testing")
st.caption(
    "File diproses sementara selama sesi aplikasi. Semua file memakai satu pipeline Generic. "
    "Klasifikasi tipe kolom adalah rekomendasi awal yang harus ditinjau sebelum proses dijalankan."
)

TYPE_LABEL_TO_KEY = {label: key for key, label in COLUMN_TYPE_LABELS.items()}
TYPE_KEY_TO_LABEL = COLUMN_TYPE_LABELS
PRESERVE_MODE_LABELS = {
    "Pertahankan semua nilai asli": "all",
    "Pertahankan nilai asli hanya ketika berubah": "changed",
    "Simpan nilai asli hanya pada Data Asli dan Audit": "audit_only",
}


@st.cache_data(show_spinner=False)
def get_sheet_names_cached(file_bytes: bytes) -> list[str]:
    return get_excel_sheet_names(file_bytes)


@st.cache_data(show_spinner=False)
def read_preview_cached(
    file_bytes: bytes,
    sheet_name: str,
    header_row: int,
    nrows: int = 10,
) -> pd.DataFrame:
    return read_preview_data(file_bytes, sheet_name, header_row, nrows=nrows)


@st.cache_data(show_spinner=False)
def read_sheet_cached(
    file_bytes: bytes,
    sheet_name: str,
    header_row: int,
) -> pd.DataFrame:
    with io.BytesIO(file_bytes) as buffer:
        return pd.read_excel(buffer, sheet_name=sheet_name, header=header_row)


def reset_result() -> None:
    for key in [
        "result_excel",
        "result_filename",
        "etl_result",
        "debug_error",
    ]:
        st.session_state.pop(key, None)


def file_fingerprint(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


def parse_length_list(text: str, field_name: str) -> list[int]:
    values: list[int] = []
    for item in text.split(","):
        item = item.strip()
        if not item:
            continue
        try:
            value = int(item)
        except ValueError as exc:
            raise ValueError(f"{field_name} harus berisi angka dipisahkan koma.") from exc
        if value <= 0:
            raise ValueError(f"{field_name} harus berisi angka positif.")
        values.append(value)
    return sorted(set(values))


def safe_dataframe(df: pd.DataFrame, **kwargs: Any) -> None:
    try:
        st.dataframe(df, **kwargs)
    except (OverflowError, TypeError, ValueError):
        st.dataframe(df.astype("string"), **kwargs)


def concise_error_message(exc: Exception) -> str:
    if isinstance(exc, MemoryError):
        return "Memori tidak cukup. Kurangi jumlah kolom atau ukuran file sebelum menjalankan ulang."
    if isinstance(exc, PermissionError):
        return "File output tidak dapat ditulis. Pastikan file Excel hasil tidak sedang dibuka."
    if isinstance(exc, KeyError):
        return f"Kolom yang dibutuhkan tidak ditemukan. Periksa mapping kolom. Detail: {exc}"
    if isinstance(exc, ValueError):
        return f"Konfigurasi atau struktur data tidak valid. Detail: {exc}"
    return "Pipeline gagal dijalankan. Periksa struktur file, mapping kolom, dan aturan cleaning."


def extract_metric(summary: pd.DataFrame, metric: str, default: Any = 0) -> Any:
    if summary.empty or "METRIC" not in summary.columns or "VALUE" not in summary.columns:
        return default
    match = summary.loc[summary["METRIC"] == metric, "VALUE"]
    return match.iloc[0] if not match.empty else default


# A. Upload file
st.header("Upload file")
uploaded_file = st.file_uploader(
    "Pilih file Excel",
    type=["xlsx", "xlsm", "xls"],
    help="File tidak disimpan permanen pada server.",
    on_change=reset_result,
)
if uploaded_file is None:
    st.info("Unggah file Excel untuk memulai konfigurasi.")
    st.stop()

file_bytes = uploaded_file.getvalue()
if not file_bytes:
    st.error("File upload kosong.")
    st.stop()

fingerprint = file_fingerprint(file_bytes)
file_info_cols = st.columns(3)
file_info_cols[0].metric("Nama file", uploaded_file.name)
file_info_cols[1].metric("Ukuran", f"{len(file_bytes) / 1024:.1f} KB")
file_info_cols[2].metric("Pipeline", PIPELINE_VERSION)

try:
    sheet_names = get_sheet_names_cached(file_bytes)
except Exception as exc:
    st.error(f"File Excel tidak dapat dibaca. File mungkin rusak atau formatnya tidak sesuai. Detail: {exc}")
    st.stop()
if not sheet_names:
    st.error("File Excel tidak memiliki sheet yang dapat dibaca.")
    st.stop()


# B. Pilih struktur file
st.header("Pilih struktur file")
structure_cols = st.columns([2, 1])
selected_sheet = structure_cols[0].selectbox(
    "Nama sheet",
    options=sheet_names,
    key="selected_sheet",
    on_change=reset_result,
)
header_row = int(
    structure_cols[1].number_input(
        "Baris header, dimulai dari 0",
        min_value=0,
        value=0,
        step=1,
        key="header_row",
        on_change=reset_result,
    )
)

try:
    preview_df = read_preview_cached(file_bytes, selected_sheet, header_row, nrows=10)
    sheet_df = read_sheet_cached(file_bytes, selected_sheet, header_row)
except MemoryError:
    st.error("Memori tidak cukup untuk membaca sheet. Kurangi ukuran file atau pilih sheet yang lebih kecil.")
    st.stop()
except Exception as exc:
    st.error(
        "Sheet atau baris header tidak dapat dibaca. Periksa apakah nomor header menunjuk baris nama kolom. "
        f"Detail: {exc}"
    )
    st.stop()

if len(sheet_df.columns) == 0:
    st.error("Baris header tidak menghasilkan nama kolom.")
    st.stop()
if sheet_df.empty:
    st.error("Sheet tidak memiliki baris data setelah header.")
    st.stop()

available_columns = [str(column) for column in sheet_df.columns]
st.session_state["uploaded_file_fingerprint"] = fingerprint
st.session_state["preview_df"] = preview_df
st.session_state["sheet_summary_df"] = sheet_df
invalid_header_columns = [
    column for column in available_columns
    if not column.strip() or column.casefold().startswith("unnamed:")
]
if invalid_header_columns:
    st.warning(
        "Header menghasilkan kolom kosong atau Unnamed: " + ", ".join(invalid_header_columns)
        + ". Perbaiki header atau jangan pilih kolom tersebut."
    )

preview_tabs = st.tabs(["Preview data", "Tipe data awal", "Nilai kosong awal"])
with preview_tabs[0]:
    safe_dataframe(preview_df, use_container_width=True, hide_index=True)
with preview_tabs[1]:
    dtype_df = pd.DataFrame({
        "Kolom": available_columns,
        "Tipe Data Awal": [str(sheet_df[column].dtype) for column in sheet_df.columns],
    })
    safe_dataframe(dtype_df, use_container_width=True, hide_index=True)
with preview_tabs[2]:
    missing_df = pd.DataFrame({
        "Kolom": available_columns,
        "Jumlah Kosong Asli": [int(sheet_df[column].isna().sum()) for column in sheet_df.columns],
    })
    safe_dataframe(missing_df, use_container_width=True, hide_index=True)

structure_summary = st.columns(4)
structure_summary[0].metric("Sheet", selected_sheet)
structure_summary[1].metric("Header row", header_row)
structure_summary[2].metric("Jumlah baris", len(sheet_df))
structure_summary[3].metric("Jumlah kolom", len(sheet_df.columns))


pipeline_mode = "Generic"

# D. Pilih dan mapping kolom
st.header("Pilih dan mapping kolom")
mapping_context = f"{fingerprint}|{selected_sheet}|{header_row}|Generic"
if st.session_state.get("mapping_context") != mapping_context:
    recommendation_df = build_mapping_recommendations(available_columns, "Generic")
    recommendation_df["Rekomendasi Tipe"] = recommendation_df["Rekomendasi Tipe"].map(TYPE_KEY_TO_LABEL)
    recommendation_df["Tipe Kolom"] = recommendation_df["Tipe Kolom"].map(TYPE_KEY_TO_LABEL)
    st.session_state["mapping_context"] = mapping_context
    st.session_state["mapping_seed"] = recommendation_df
    reset_result()

mapping_df = st.data_editor(
    st.session_state["mapping_seed"],
    key=f"mapping_editor_{hashlib.md5(mapping_context.encode()).hexdigest()}",
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    disabled=["Kolom Sumber", "Rekomendasi Tipe"],
    on_change=reset_result,
    column_config={
        "Analisis": st.column_config.CheckboxColumn("Analisis"),
        "Kolom Sumber": st.column_config.TextColumn("Kolom Sumber"),
        "Nama Kanonik": st.column_config.TextColumn(
            "Nama Kanonik",
            help="Nama output setelah mapping. Ubah jika rekomendasi tidak sesuai.",
        ),
        "Rekomendasi Tipe": st.column_config.TextColumn("Rekomendasi Tipe"),
        "Tipe Kolom": st.column_config.SelectboxColumn(
            "Tipe Kolom",
            options=list(TYPE_LABEL_TO_KEY.keys()),
            required=True,
        ),
    },
)

selected_mapping_df = mapping_df[mapping_df["Analisis"].fillna(False)].copy()
selected_columns = selected_mapping_df["Kolom Sumber"].astype(str).tolist()
column_mapping = {
    str(row["Kolom Sumber"]): TYPE_LABEL_TO_KEY[str(row["Tipe Kolom"])]
    for _, row in selected_mapping_df.iterrows()
}
canonical_mapping = {
    str(row["Kolom Sumber"]): str(row["Nama Kanonik"]).strip()
    for _, row in selected_mapping_df.iterrows()
}

if not selected_columns:
    st.warning("Minimal satu kolom harus dipilih untuk dianalisis.")

with st.expander("Aturan kolom", expanded=True):
    required_columns = st.multiselect(
        "Kolom yang tidak boleh kosong",
        options=selected_columns,
        default=[c for c in st.session_state.get("required_columns", []) if c in selected_columns],
        key="required_columns",
        help="Kolom ini dilaporkan sebagai pelanggaran jika kosong atau placeholder.",
        on_change=reset_result,
    )
    duplicate_key_columns = st.multiselect(
        "Kolom pembanding duplikat",
        options=selected_columns,
        default=[c for c in st.session_state.get("duplicate_key_columns", []) if c in selected_columns],
        key="duplicate_key_columns",
        help=(
            "Logika menggunakan OR. Nilai berulang pada salah satu kolom terpilih sudah dianggap "
            "sebagai temuan duplikat dan alasan dicatat per kolom serta per nilai."
        ),
        on_change=reset_result,
    )


# E. Aturan cleaning
st.header("Aturan cleaning")
cleaning_col1, cleaning_col2, cleaning_col3 = st.columns(3)
with cleaning_col1:
    preserve_label = st.selectbox(
        "Penyimpanan nilai asli",
        options=list(PRESERVE_MODE_LABELS.keys()),
        index=0,
        on_change=reset_result,
    )
    create_audit = st.checkbox("Buat audit perubahan", value=True, on_change=reset_result)
    create_detail_reports = st.checkbox("Buat laporan detail", value=True, on_change=reset_result)
with cleaning_col2:
    remove_empty_required = st.checkbox(
        "Hapus baris jika kolom wajib kosong",
        value=False,
        on_change=reset_result,
    )
    remove_invalid_phone = st.checkbox(
        "Hapus baris dengan nomor telepon invalid",
        value=False,
        on_change=reset_result,
    )
    remove_invalid_email = st.checkbox(
        "Hapus baris dengan email invalid",
        value=False,
        on_change=reset_result,
    )
with cleaning_col3:
    keep_other_anomalies = st.checkbox(
        "Pertahankan data anomali lainnya",
        value=True,
        on_change=reset_result,
    )
    remove_duplicates = st.checkbox(
        "Hapus data duplikat",
        value=False,
        on_change=reset_result,
    )
    duplicate_keep_label = st.selectbox(
        "Saat menghapus duplikat",
        options=["Pertahankan baris pertama", "Pertahankan baris terakhir", "Hapus seluruh kelompok"],
        disabled=not remove_duplicates,
        on_change=reset_result,
    )

duplicate_keep_map = {
    "Pertahankan baris pertama": "first",
    "Pertahankan baris terakhir": "last",
    "Hapus seluruh kelompok": "none",
}

with st.expander("Aturan placeholder dan batas validasi", expanded=False):
    placeholder_text = st.text_area(
        "Placeholder, satu nilai per baris",
        value="\n".join(DEFAULT_PLACEHOLDERS),
        height=220,
        on_change=reset_result,
    )
    limit_cols = st.columns(4)
    min_year = int(limit_cols[0].number_input("Tahun minimum", min_value=1800, value=2000, step=1, on_change=reset_result))
    max_year = int(limit_cols[1].number_input("Tahun maksimum", min_value=1800, value=pd.Timestamp.today().year, step=1, on_change=reset_result))
    policy_min_length = int(limit_cols[2].number_input("Panjang minimum nomor polis", min_value=1, value=3, step=1, on_change=reset_result))
    report_min_length = int(limit_cols[3].number_input("Panjang minimum nomor pelaporan", min_value=1, value=3, step=1, on_change=reset_result))
    identity_cols = st.columns(3)
    nik_lengths_text = identity_cols[0].text_input("Panjang NIK", value="16", on_change=reset_result)
    npwp_lengths_text = identity_cols[1].text_input("Panjang NPWP", value="15,16", on_change=reset_result)
    identity_lengths_text = identity_cols[2].text_input(
        "Panjang identitas umum, kosong berarti bebas",
        value="",
        on_change=reset_result,
    )

placeholder_values = [line.strip() for line in placeholder_text.splitlines() if line.strip()]

try:
    nik_lengths = parse_length_list(nik_lengths_text, "Panjang NIK")
    npwp_lengths = parse_length_list(npwp_lengths_text, "Panjang NPWP")
    identity_lengths = parse_length_list(identity_lengths_text, "Panjang identitas umum")
    length_parse_error = None
except ValueError as exc:
    nik_lengths, npwp_lengths, identity_lengths = [], [], []
    length_parse_error = str(exc)

# F. Pilih sheet output
st.header("Pilih sheet output")
output_labels = {label: key for key, label in REPORT_REGISTRY.items()}
default_output_labels = [REPORT_REGISTRY[key] for key in DEFAULT_OUTPUT_SHEETS]
selected_output_labels = st.multiselect(
    "Sheet yang ditulis ke file Excel",
    options=list(output_labels.keys()),
    default=default_output_labels,
    help="Analisis inti tetap dijalankan. Pilihan ini hanya menentukan laporan yang ditulis ke Excel.",
    on_change=reset_result,
)
output_sheets = [output_labels[label] for label in selected_output_labels]

if not output_sheets:
    st.error("Minimal satu sheet output harus dipilih.")
if "duplicates" in output_sheets and not duplicate_key_columns:
    st.warning(
        "Sheet Data Duplikat dipilih, tetapi kolom pembanding duplikat belum ditentukan. "
        "Sheet tetap dibuat dan dapat kosong; laporan duplikat legacy tetap tersedia bila dipilih."
    )
if "removed" in output_sheets and not any([
    remove_empty_required, remove_invalid_phone, remove_invalid_email,
    not keep_other_anomalies, remove_duplicates,
]):
    st.info("Data Dihapus dipilih, tetapi tidak ada aturan penghapusan aktif. Sheet akan kosong.")



config_for_validation = ETLConfig(
    input_file="temporary_input.xlsx",
    output_file="temporary_output.xlsx",
    input_filename=uploaded_file.name,
    sheet_name=selected_sheet,
    header=header_row,
    selected_columns=selected_columns,
    pipeline_mode=pipeline_mode,
    column_mapping=column_mapping,
    canonical_mapping=canonical_mapping,
    required_columns=required_columns,
    duplicate_key_columns=duplicate_key_columns,
    placeholder_values=placeholder_values,
    preserve_original_mode=PRESERVE_MODE_LABELS[preserve_label],
    remove_empty_required=remove_empty_required,
    remove_invalid_phone=remove_invalid_phone,
    remove_invalid_email=remove_invalid_email,
    remove_invalid_other=not keep_other_anomalies,
    remove_duplicates=remove_duplicates,
    duplicate_keep=duplicate_keep_map[duplicate_keep_label],
    create_audit=create_audit,
    create_detail_reports=create_detail_reports,
    min_year=min_year,
    max_year=max_year,
    nik_lengths=nik_lengths,
    npwp_lengths=npwp_lengths,
    identity_lengths=identity_lengths,
    policy_min_length=policy_min_length,
    report_number_min_length=report_min_length,
    output_sheets=output_sheets,
)
st.session_state["current_etl_config"] = config_for_validation
validation_errors, validation_warnings = validate_etl_config(config_for_validation, available_columns)
if length_parse_error:
    validation_errors.append(length_parse_error)


# G. Preview konfigurasi
st.header("Preview konfigurasi")
configuration_preview = {
    "file": uploaded_file.name,
    "sheet": selected_sheet,
    "header_row": header_row,
    "row_count": len(sheet_df),
    "selected_column_count": len(selected_columns),
    "pipeline_mode": "Generic",
    "required_columns": required_columns,
    "duplicate_key_columns": duplicate_key_columns,
    "remove_empty_required": remove_empty_required,
    "remove_invalid_phone": remove_invalid_phone,
    "remove_invalid_email": remove_invalid_email,
    "remove_invalid_other": not keep_other_anomalies,
    "remove_duplicates": remove_duplicates,
    "preserve_original_mode": PRESERVE_MODE_LABELS[preserve_label],
    "output_sheets": [REPORT_REGISTRY[key] for key in output_sheets],
}
preview_config_cols = st.columns([1, 2])
with preview_config_cols[0]:
    st.json(configuration_preview)
with preview_config_cols[1]:
    safe_dataframe(
        selected_mapping_df[[
            "Kolom Sumber", "Nama Kanonik", "Tipe Kolom"
        ]],
        use_container_width=True,
        hide_index=True,
    )

for warning in validation_warnings:
    st.warning(warning)
for error in validation_errors:
    st.error(error)

configuration_valid = not validation_errors


# H. Proses
st.header("Proses")
run_button = st.button(
    "Jalankan Pipeline",
    type="primary",
    use_container_width=True,
    disabled=not configuration_valid,
)

if run_button:
    input_suffix = Path(uploaded_file.name).suffix.lower() or ".xlsx"
    safe_stem = Path(sanitize_filename(uploaded_file.name, "input.xlsx")).stem
    output_filename = sanitize_filename(f"hasil_data_quality_{safe_stem}.xlsx")
    progress_bar = st.progress(0)
    status_area = st.empty()

    def progress_callback(percent: int, message: str) -> None:
        progress_bar.progress(percent)
        status_area.write(message)

    try:
        with tempfile.TemporaryDirectory(prefix="askrindo_etl_") as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / sanitize_filename(f"input{input_suffix}")
            output_path = temp_path / output_filename
            input_path.write_bytes(file_bytes)

            final_config = ETLConfig(
                **{
                    **config_for_validation.__dict__,
                    "input_file": input_path,
                    "output_file": output_path,
                    "progress_callback": progress_callback,
                }
            )
            result = run_etl(final_config)
            if not output_path.exists() or output_path.stat().st_size == 0:
                raise RuntimeError("Pipeline tidak menghasilkan file Excel.")
            st.session_state["result_excel"] = output_path.read_bytes()
            st.session_state["result_filename"] = output_filename
            st.session_state["etl_result"] = result
            st.session_state.pop("debug_error", None)
        st.success("Pipeline selesai dijalankan.")
    except Exception as exc:
        st.session_state["debug_error"] = traceback.format_exc()
        st.error(concise_error_message(exc))


# I. Hasil
if "result_excel" in st.session_state and "etl_result" in st.session_state:
    st.header("I. Hasil")
    result = st.session_state["etl_result"]
    summary = result["summary"]

    metric_columns = st.columns(6)
    metric_columns[0].metric("Baris awal", extract_metric(summary, "Initial Rows"))
    metric_columns[1].metric("Baris bersih", extract_metric(summary, "Clean Rows"))
    metric_columns[2].metric("Baris dihapus", extract_metric(summary, "Removed Rows"))
    metric_columns[3].metric("Baris invalid", extract_metric(summary, "Invalid Rows"))
    metric_columns[4].metric("Baris duplikat", extract_metric(summary, "Duplicate Rows"))
    completeness_value = extract_metric(summary, "Overall Completeness (%)", None)
    metric_columns[5].metric(
        "Completeness",
        "N/A" if completeness_value is None or pd.isna(completeness_value) else f"{float(completeness_value):.2f}%",
    )

    report_tabs = st.tabs([
        "Data Bersih",
        "Completeness",
        "Validity",
        "Data Invalid",
        "Data Duplikat",
        "Audit",
    ])
    with report_tabs[0]:
        safe_dataframe(result["clean"].head(200), use_container_width=True, hide_index=True)
    with report_tabs[1]:
        completeness_df = result["reports"]["Completeness Kolom"]
        if not completeness_df.empty:
            st.bar_chart(completeness_df.set_index("COLUMN")["COMPLETENESS_PCT"])
        safe_dataframe(completeness_df, use_container_width=True, hide_index=True)
    with report_tabs[2]:
        validity_df = result["reports"]["Validity Kolom"]
        if not validity_df.empty:
            st.bar_chart(validity_df.set_index("COLUMN")["VALIDITY_PCT"])
        safe_dataframe(validity_df, use_container_width=True, hide_index=True)
    with report_tabs[3]:
        invalid_df = result["reports"]["Data Tidak Valid"]
        invalid_rows = int(invalid_df["ROW_ID"].nunique()) if not invalid_df.empty and "ROW_ID" in invalid_df.columns else 0
        st.caption(
            f"{invalid_rows} baris memiliki minimal satu nilai invalid; "
            f"terdapat {len(invalid_df)} temuan invalid pada tingkat sel."
        )
        safe_dataframe(invalid_df.head(1000), use_container_width=True, hide_index=True)
    with report_tabs[4]:
        duplicate_df = result["reports"]["Data Duplikat"]
        duplicate_rows = int(duplicate_df["ROW_ID"].nunique()) if not duplicate_df.empty and "ROW_ID" in duplicate_df.columns else 0
        st.caption(
            f"{duplicate_rows} baris terlibat dalam duplikat; "
            f"terdapat {len(duplicate_df)} temuan karena satu baris dapat memenuhi beberapa kondisi duplikat."
        )
        safe_dataframe(duplicate_df.head(1000), use_container_width=True, hide_index=True)
    with report_tabs[5]:
        safe_dataframe(result["reports"]["Audit Perubahan"].head(1000), use_container_width=True, hide_index=True)

    st.download_button(
        "Unduh hasil Excel",
        data=st.session_state["result_excel"],
        file_name=st.session_state["result_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    if result.get("warnings"):
        with st.expander("Peringatan proses", expanded=False):
            for warning in result["warnings"]:
                st.write(warning)
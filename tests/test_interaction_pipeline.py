from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
import pytest

from interaction_pipeline import (
    InteractionConfig,
    company_match_key,
    export_interaction_excel,
    normalize_company_name,
    normalize_person_name,
    person_name_match_key,
    run_interaction_pipeline,
)


def make_config(**overrides):
    values = dict(
        master_sheet="Master",
        master_header=0,
        master_row_id_column="ROW_ID",
        master_customer_name_column="Nama",
        master_company_name_column="Perusahaan",
        master_phone_columns=["HP"],
        master_email_columns=["Email"],
        contact_sheet="Contact",
        contact_header=0,
        contact_date_column="Tanggal",
        contact_report_number_column="No Laporan",
        contact_customer_name_column="Nama",
        contact_company_name_column="Perusahaan",
        contact_phone_columns=["HP"],
        contact_email_columns=["Email"],
        contact_channel_column="Saluran",
        contact_category_column="Kategori",
        match_phone=True,
        match_email=True,
        match_name=True,
        match_company=True,
        allow_name_company_fallback=True,
        interaction_definition="report_number",
    )
    values.update(overrides)
    return InteractionConfig(**values)


def master_df():
    return pd.DataFrame({
        "ROW_ID": ["R1", "R2", "R3"],
        "Nama": ["Budi Santoso", "Siti Aminah", "Andi"],
        "Perusahaan": ["PT ABC Indonesia", "CV Maju", "XYZ"],
        "HP": ["081234567890", "081298765432", "081211111111"],
        "Email": ["budi@example.com", "siti@example.com", "andi@example.com"],
    })


def contact_df():
    return pd.DataFrame({
        "Tanggal": ["01/01/2026", "15/01/2026", "20/01/2026", "01/02/2026"],
        "No Laporan": ["L1", "L2", "L3", "L4"],
        "Nama": ["Budi Santoso", "Budi Santoso", "Siti Aminah", "Unknown"],
        "Perusahaan": ["ABC Indonesia", "PT ABC Indonesia", "CV Maju", "Unknown Co"],
        "HP": ["081234567890", "081234567890", "", "089999999999"],
        "Email": ["budi@example.com", "", "siti@example.com", "unknown@example.com"],
        "Saluran": ["WA", "Telp", "Email", "WA"],
        "Kategori": ["A", "B", "A", "C"],
    })


def test_name_normalization():
    assert normalize_person_name("  bUDI   sANTOSO ") == "Budi Santoso"
    assert person_name_match_key("Budi, Santoso") == "budi santoso"


def test_company_normalization_without_guessing_legal_entity():
    assert normalize_company_name("P.T. ABC INDONESIA") == "PT ABC Indonesia"
    assert normalize_company_name("ABC INDONESIA") == "ABC Indonesia"
    assert company_match_key("PT ABC Indonesia") == "abc indonesia"
    assert company_match_key("CV ABC Indonesia") == "abc indonesia"


def test_phone_email_phone_email_and_not_found():
    result = run_interaction_pipeline(master_df(), contact_df(), make_config())["result"]
    assert result["MATCH_METHOD"].tolist() == ["PHONE_EMAIL", "PHONE", "EMAIL", "NONE"]
    assert result["MATCH_STATUS"].tolist() == ["MATCHED", "MATCHED", "MATCHED", "NOT_FOUND"]


def test_name_company_fallback():
    contact = contact_df().iloc[[0]].copy()
    contact["HP"] = ""
    contact["Email"] = ""
    result = run_interaction_pipeline(master_df(), contact, make_config())["result"].iloc[0]
    assert result["MATCH_STATUS"] == "MATCHED"
    assert result["MATCH_METHOD"] == "NAME_COMPANY"
    assert result["MATCHED_CUSTOMER_ROW_ID"] == "R1"


def test_customer_without_contact_is_retained():
    result = run_interaction_pipeline(master_df(), contact_df().iloc[:1], make_config())
    summary = result["customer_summary"].set_index("MASTER_ROW_ID")
    assert len(summary) == 3
    assert int(summary.loc["R2", "TOTAL_CONTACTS"]) == 0
    assert bool(summary.loc["R2", "EVER_CONTACTED"]) is False


def test_unique_report_count_and_sequence():
    contact = contact_df().iloc[:2].copy()
    contact.loc[contact.index[1], "No Laporan"] = "L1"
    result = run_interaction_pipeline(master_df(), contact, make_config())
    summary = result["customer_summary"].set_index("MASTER_ROW_ID")
    assert int(summary.loc["R1", "TOTAL_CONTACTS"]) == 1
    assert result["detail"]["CONTACT_SEQUENCE"].tolist() == [1, 2]


def test_duplicate_master_row_id_rejected():
    master = master_df()
    master.loc[1, "ROW_ID"] = "R1"
    with pytest.raises(ValueError, match="ROW_ID master harus unik"):
        run_interaction_pipeline(master, contact_df(), make_config())


def test_excel_export():
    result = run_interaction_pipeline(master_df(), contact_df(), make_config())
    with TemporaryDirectory() as temp_dir:
        output = export_interaction_excel(result, Path(temp_dir) / "hasil.xlsx")
        assert output.exists()
        assert output.stat().st_size > 0

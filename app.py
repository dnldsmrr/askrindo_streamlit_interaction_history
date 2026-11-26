from __future__ import annotations

import streamlit as st

st.set_page_config(page_title="ASKRINDO Data Tools", layout="wide")

st.title("Customer Care Data Tools")
st.write(
    "Aplikasi dipisahkan menjadi dua halaman karena analisis kualitas data dan "
    "identifikasi customer memiliki tujuan, input, dan output yang berbeda."
)

col1, col2 = st.columns(2)
with col1:
    st.subheader("Analisis Dataset")
    st.write(
        "Satu file input untuk normalisasi, validasi, completeness, duplikat, "
        "relasi satu-ke-banyak, cleaning, audit, dan ekspor Excel."
    )
    st.page_link("pages/1_Analisis_Dataset.py", label="Buka Analisis Dataset")

with col2:
    st.subheader("Riwayat Interaksi Pelanggan")
    st.write(
        "Dua file input: master pelanggan dan data Contact Centre. Tujuan halaman ini adalah "
        "mencatat riwayat interaksi pelanggan, mencocokkan kontak dengan ROW_ID master, dan "
        "menghasilkan ringkasan interaksi tanpa melakukan analisis kualitas dataset umum."
    )
    st.page_link("pages/2_Riwayat_Interaksi_Pelanggan.py", label="Buka Riwayat Interaksi Pelanggan")

st.divider()
st.caption(
    "File upload diproses di memori atau folder sementara. Tidak ada penyimpanan permanen pada server."
)

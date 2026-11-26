# ASKRINDO Data Tools

Aplikasi terdiri dari dua halaman:

1. Analisis Dataset: pipeline data quality Generic.
2. Riwayat Interaksi Pelanggan: matching Contact Centre ke ROW_ID master dan perhitungan riwayat interaksi.

## Menjalankan

```powershell
py -3.14 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements_streamlit_askrindo.txt
streamlit run app.py
```

File upload dan output diproses melalui memori atau direktori sementara.

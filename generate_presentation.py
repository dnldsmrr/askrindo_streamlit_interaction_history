from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt

OUTPUT_FILE = Path(r"D:\PKL ASKRINDO\ASKRINDO_Project_Presentation.pptx")

PROJECT_TITLE = "ASKRINDO Data Tools"
PROJECT_SUBTITLE = "Streamlit application for data quality analysis and customer interaction history"

SLIDES = [
    {
        "title": "Ringkasan Proyek",
        "bullets": [
            "Aplikasi Streamlit untuk analisis kualitas data dan riwayat interaksi pelanggan.",
            "Dua halaman utama: Data Quality Testing dan Riwayat Interaksi Pelanggan.",
            "Menangani upload Excel, normalisasi, validasi, matching, dan ekspor hasil.",
        ],
    },
    {
        "title": "Halaman 1: Data Quality Testing",
        "bullets": [
            "Upload file Excel kemudian pilih sheet dan header.",
            "Preview data, tipe awal, dan jumlah nilai kosong.",
            "Mapping kolom ke tipe data kanonik dan aturan cleaning.",
            "Output: file Excel dengan sheet kualitas data, validitas, duplikat, dan audit.",
        ],
    },
    {
        "title": "Halaman 2: Riwayat Interaksi Pelanggan",
        "bullets": [
            "Upload master pelanggan dan data Contact Centre.",
            "Mapping kolom ROW_ID, telepon, email, nama, perusahaan, saluran, kategori, dll.",
            "Matching multi-jenis: telepon, email, dan fallback nama+perusahaan.",
            "Ringkasan per pelanggan, detail kontak, ambiguous, not found, dan audit matching.",
        ],
    },
    {
        "title": "Pipeline dan Logika Utama",
        "bullets": [
            "`pipeline_etl_askrindo.py` untuk data quality testing dan normalisasi generic.",
            "`interaction_pipeline.py` untuk persiapan master, normalisasi kontak, dan matching.",
            "Validasi konfigurasi sebelum proses dijalankan untuk memastikan mapping lengkap.",
            "Output Excel dibuat dengan format sheet yang mudah dibaca, difilter, dan ditebalkan header-nya.",
        ],
    },
    {
        "title": "Sorotan Fitur Teknis",
        "bullets": [
            "Normalisasi telepon, email, nama pelanggan, dan nama perusahaan sebelum matching.",
            "Matching bertahap: PHONE_EMAIL -> PHONE -> EMAIL -> NAME_COMPANY -> PHONE_COMPANY -> EMAIL_COMPANY.",
            "Fallback nama+perusahaan dipakai hanya saat telepon dan email tidak cukup.",
            "Validasi rule: placeholder kustom, panjang NIK/NPWP, duplikat, dan baris required.",
        ],
    },
    {
        "title": "Contoh Output Hasil",
        "bullets": [
            "Ringkasan Proses: metrik pipeline, jumlah matched/ambiguous/not_found, completeness.",
            "Master Normalisasi & Contact Centre Normalisasi: semua sumber kontak disatukan ke satu kolom telepon/email.",
            "Hasil Matching: status, candidate list, row ID target, alasan matching, dan skor.",
            "Ringkasan per Pelanggan: total kontak, tanggal pertama/terakhir, channel, kategori, dan detail riwayat.",
        ],
    },
    {
        "title": "Visualisasi Alur Data",
        "bullets": [
            "1. Upload file Excel -> Baca sheet, pilih header, dan preview data.",
            "2. Normalisasi dan validasi data: telepon, email, nama, perusahaan, placeholder, dan duplikat.",
            "3. Matching Contact Centre ke master pelanggan secara bertahap.",
            "4. Bangun summary pelanggan dan detail riwayat untuk audit dan analisis.",
            "5. Ekspor hasil ke Excel dengan sheet terpisah untuk review dan dokumentasi.",
        ],
    },
    {
        "title": "Struktur Proyek",
        "bullets": [
            "app.py: halaman utama dengan navigasi ke dua modul.",
            "pages/1_Analisis_Dataset.py: antarmuka Data Quality Testing.",
            "pages/2_Riwayat_Interaksi_Pelanggan.py: antarmuka matching interaksi pelanggan.",
            "pipeline_etl_askrindo.py: logika ETL, validasi, normalisasi, dan laporan.",
            "interaction_pipeline.py: logika matching Contact Centre ke master pelanggan.",
            "generate_presentation.py: membuat materi PPT secara otomatis dari struktur proyek.",
        ],
    },
    {
        "title": "Output dan Export",
        "bullets": [
            "Hasil export ke Excel: ringkasan proses, konfigurasi, data normalisasi, dan audit.",
            "Riwayat Interaksi Pelanggan menghasilkan sheet detail, ringkasan pelanggan, dan status matching.",
            "File tidak disimpan permanen di server; semua upload diproses sementara.",
        ],
    },
    {
        "title": "Instruksi Menjalankan Aplikasi",
        "bullets": [
            "Buat virtual environment: py -3.14 -m venv .venv.",
            "Aktifkan dan install dependensi: pip install -r requirements_streamlit_askrindo.txt.",
            "Jalankan Streamlit: streamlit run app.py.",
            "Buka aplikasi dan pilih halaman sesuai kebutuhan analisis.",
        ],
    },
    {
        "title": "Manfaat dan Keunggulan",
        "bullets": [
            "Mempercepat validasi dan cleaning data Excel secara interaktif.",
            "Menjadikan data kontak pelanggan terhubung ke master dengan lebih akurat.",
            "Menghasilkan laporan lengkap untuk audit internal dan pemantauan kualitas data.",
        ],
    },
]


def add_bullet_slide(prs: Presentation, title: str, bullets: list[str]) -> None:
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    body = slide.shapes.placeholders[1].text_frame
    body.text = bullets[0]
    for bullet in bullets[1:]:
        p = body.add_paragraph()
        p.text = bullet
        p.level = 0
        p.font.size = Pt(18)


def create_presentation(output_path: Path) -> None:
    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = PROJECT_TITLE
    subtitle = slide.placeholders[1]
    subtitle.text = PROJECT_SUBTITLE

    for slide_data in SLIDES:
        add_bullet_slide(prs, slide_data["title"], slide_data["bullets"])

    prs.save(output_path)
    print(f"Generated presentation: {output_path}")


if __name__ == "__main__":
    create_presentation(OUTPUT_FILE)

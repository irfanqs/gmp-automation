# GMP Automation System

Aplikasi web lokal untuk mengubah PDF hasil pengukuran lingkungan pabrik farmasi menjadi laporan Microsoft Excel yang rapi, lengkap dengan data terstruktur, nilai rata-rata, penanda batas, dan grafik.

## Fitur

- Memproses banyak PDF sekaligus untuk AHU yang berbeda.
- Mendukung ekstraksi dengan Claude API atau DeepSeek-OCR melalui backend Kaggle.
- Menghasilkan workbook Excel per jenis pengujian dengan sheet data, tabel ringkasan, dan grafik.
- Menandai nilai di luar batas dengan warna merah.
- Mengelompokkan hasil berdasarkan AHU dan semester pengukuran.

## Jenis Pengujian

| Kode | Pengujian | Berkas Excel |
| --- | --- | --- |
| A | Airborne Particle Test | `Airborne_Particle_Test_Result_and_Graph.xlsx` |
| B | Air Velocity Test | `Air_Velocity_Test_Result_and_Graph.xlsx` |
| C | Air Change Rate Test | `Air_Change_Rate_Test_Result_and_Graph.xlsx` |
| D | HEPA Filter Test | `HEPA_Filter_Test_Result_and_Graph.xlsx` |
| E | Airflow Pattern Test | `Airflow_Pattern_Test_Result_and_Graph.xlsx` |

## Kebutuhan Sistem

- Python 3.10 atau lebih baru.
- Poppler, untuk mengubah halaman PDF menjadi gambar.
- Koneksi internet untuk OCR.
- Salah satu backend OCR berikut:
  - Anthropic API key untuk Claude API.
  - Notebook Kaggle DeepSeek-OCR yang sedang berjalan dan URL ngrok-nya.

Instal Poppler:

```bash
# macOS
brew install poppler

# Ubuntu/Debian
sudo apt install poppler-utils
```

Di Windows, unduh Poppler dari [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases), ekstrak misalnya ke `C:\poppler`, lalu tambahkan `C:\poppler\Library\bin` ke `PATH`.

## Instalasi

```bash
git clone <URL_REPOSITORI>
cd gmp_automation
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Untuk Windows, aktifkan virtual environment dengan:

```bat
.venv\Scripts\activate
```

## Menjalankan Aplikasi

```bash
python app.py
```

Atau gunakan skrip:

```bash
# macOS/Linux
bash START_LINUX.sh

# Windows
START_WINDOWS.bat
```

Buka `http://localhost:5001` di browser.

## Cara Menggunakan

1. Pilih backend OCR.
2. Masukkan Anthropic API key, atau URL endpoint ngrok dari DeepSeek-OCR Kaggle.
3. Pilih satu jenis pengujian.
4. Unggah satu atau beberapa PDF dengan jenis pengujian yang sama.
5. Klik tombol pembuatan Excel dan unduh berkas hasilnya.

Setiap PDF sebaiknya memuat satu AHU untuk satu semester. Kualitas hasil bergantung pada keterbacaan scan PDF.

## DeepSeek-OCR dengan Kaggle

DeepSeek-OCR adalah alternatif tanpa biaya API Anthropic. Jalankan `deepseek_ocr/kaggle_server.ipynb` di Kaggle dengan GPU dan internet aktif.

1. Buat notebook Kaggle, lalu impor `deepseek_ocr/kaggle_server.ipynb`.
2. Atur accelerator ke GPU dan aktifkan Internet.
3. Tambahkan Kaggle Secret bernama `NGROK_AUTHTOKEN`.
4. Jalankan semua sel notebook.
5. Salin URL `https://*.ngrok-free.app` yang ditampilkan notebook ke aplikasi web.

Sesi Kaggle dan URL ngrok bersifat sementara. Parser DeepSeek bergantung pada format dokumen GMP yang dikenali; gunakan `debug_ocr.py` untuk melihat teks OCR mentah apabila hasil ekstraksi perlu diperiksa.

```bash
python debug_ocr.py <path_pdf> <ngrok_url>
```

## Struktur Proyek

```text
gmp_automation/
├── app.py                 # Aplikasi Flask dan endpoint upload/download
├── config.py              # Konfigurasi batas nilai dan jenis pengujian
├── ocr_engine.py          # Ekstraksi PDF melalui Claude API
├── deepseek_ocr/          # Klien, parser, dan notebook backend DeepSeek-OCR
├── excel_generator.py     # Pembuatan laporan Excel dan grafik
├── templates/index.html   # Antarmuka web
├── boilerplate/           # Contoh/template Excel
├── uploads/               # Penyimpanan PDF sementara saat diproses
├── outputs/               # Excel hasil proses
├── debug_ocr.py           # Alat inspeksi hasil OCR DeepSeek
└── requirements.txt       # Dependensi Python
```

## Catatan

- Claude API memerlukan biaya sesuai penggunaan akun Anthropic.
- Nilai batas dan pengaturan Excel berada di `config.py`.
- Aplikasi membatasi total unggahan permintaan hingga 100 MB.
- Berkas unggahan yang berhasil diproses dihapus setelah Excel dibuat.

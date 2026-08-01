# GMP Automation System
# Sistem Automasi Pengukuran Lingkungan Pabrik Farmasi

## 📋 Deskripsi
Sistem ini secara otomatis mengubah PDF hasil scan pengukuran lingkungan (dari vendor)
menjadi file Microsoft Excel yang rapi dan lengkap dengan tabel data, nilai rata-rata,
conditional formatting (pewarnaan), dan chart.

### 5 Jenis Pengukuran yang Didukung:
| No | Jenis Pengukuran | Input PDF | Output Excel |
|----|-----------------|-----------|-------------|
| A | Airborne Particle Test | PDF AA | Airborne_Particle_Test_Result_and_Graph.xlsx |
| B | Air Velocity Test | PDF BB | Air_Velocity_Test_Result_and_Graph.xlsx |
| C | Air Change Rate Test | PDF CC | Air_Change_Rate_Test_Result_and_Graph.xlsx |
| D | HEPA Filter Test | PDF DD | HEPA_Filter_Test_Result_and_Graph.xlsx |
| E | Airflow Pattern Test | PDF EE | Airflow_Pattern_Test_Result_and_Graph.xlsx |

---

## 🔧 Cara Install (Pertama Kali Saja)

### Langkah 1: Install Python
1. Download Python dari: https://www.python.org/downloads/
2. Saat install, **WAJIB centang** ✅ "Add Python to PATH"
3. Klik "Install Now"

### Langkah 2: Install Poppler (untuk membaca PDF)
**Windows:**
1. Download Poppler dari: https://github.com/oschwartz10612/poppler-windows/releases
2. Extract/unzip ke folder `C:\poppler`
3. Buka Settings → System → About → Advanced system settings → Environment Variables
4. Di "System Variables", klik "Path" → Edit → New
5. Tambahkan: `C:\poppler\Library\bin`
6. Klik OK semua

**Mac:**
```
brew install poppler
```

**Linux (Ubuntu/Debian):**
```
sudo apt install poppler-utils
```

### Langkah 3: Dapatkan Anthropic API Key
1. Buka https://console.anthropic.com/
2. Buat akun (jika belum punya)
3. Pergi ke menu "API Keys"
4. Klik "Create Key"
5. Copy API key yang muncul (dimulai dengan `sk-ant-...`)
6. Simpan API key ini baik-baik

### Langkah 4: Install Dependencies Python
Buka Command Prompt / Terminal, lalu jalankan:
```
cd [folder dimana file ini berada]
pip install -r requirements.txt
```

---

## 🚀 Cara Menggunakan

### Langkah 1: Jalankan Sistem
**Windows:** Klik 2x file `START_WINDOWS.bat`
**Linux/Mac:** Jalankan `bash START_LINUX.sh`

Atau buka Command Prompt / Terminal:
```
cd [folder dimana file ini berada]
python app.py
```

### Langkah 2: Buka Browser
Buka browser (Chrome/Edge) dan pergi ke: **http://localhost:5000**

### Langkah 3: Gunakan Sistem
1. **Masukkan API Key** → Paste API key Anthropic Anda
2. **Pilih Jenis Pengukuran** → Klik salah satu dari 5 jenis
3. **Upload PDF** → Drag & drop atau klik untuk memilih file PDF
   - Anda bisa upload banyak PDF sekaligus (untuk berbagai AHU)
   - Pastikan semua PDF adalah jenis pengukuran yang sama
   - Setiap PDF = 1 AHU × 1 semester
4. **Klik "Excel 자동 생성 시작"** → Tunggu proses selesai
5. **Download Excel** → Klik tombol download hijau

---

## 📁 Struktur Folder
```
gmp_automation/
├── app.py                    ← Aplikasi utama
├── config.py                  ← Konfigurasi dan konstanta
├── ocr_engine.py               ← Mesin OCR (Claude API)
├── deepseek_ocr/                ← Semua kode terkait DeepSeek-OCR (folder terpisah)
│   ├── client.py                  ← Client HTTP ke backend DeepSeek-OCR (Kaggle)
│   ├── parsers.py                 ← Parser markdown/teks OCR → JSON terstruktur
│   ├── engine.py                  ← Mesin OCR alternatif (dipanggil dari app.py)
│   └── kaggle_server.ipynb        ← Notebook backend (jalankan di Kaggle)
├── excel_generator.py          ← Generator file Excel
├── templates/
│   └── index.html              ← Tampilan web
├── uploads/                    ← Temporary (PDF upload)
├── outputs/                    ← Hasil Excel
├── requirements.txt            ← Daftar library Python
├── START_WINDOWS.bat           ← Script start (Windows)
├── START_LINUX.sh              ← Script start (Linux/Mac)
└── README.md                   ← File ini
```

---

## 🧠 Opsi Alternatif: OCR Gratis via DeepSeek-OCR (Kaggle)

Selain Claude API (berbayar), sistem ini juga mendukung backend OCR gratis menggunakan
model [DeepSeek-OCR](https://huggingface.co/deepseek-ai/DeepSeek-OCR) yang dijalankan di Kaggle (GPU gratis).

### Cara pakai:
1. Buka https://kaggle.com/code, buat notebook baru, lalu **upload/import** file
   `deepseek_ocr/kaggle_server.ipynb` dari folder ini.
2. Di notebook: **Settings → Accelerator → GPU T4 x2**, **Settings → Internet → ON**.
3. Buat akun ngrok gratis (https://dashboard.ngrok.com/get-started/your-authtoken),
   lalu tambahkan authtoken sebagai **Kaggle Secret** bernama `NGROK_AUTHTOKEN`
   (menu Add-ons → Secrets di notebook).
4. Klik **Run All**. Di output cell terakhir akan muncul URL publik seperti
   `https://xxxx.ngrok-free.app` — biarkan notebook tetap berjalan (jangan di-stop).
5. Di aplikasi web lokal, pada Langkah 1 pilih **"DeepSeek-OCR (Kaggle)"**, lalu
   tempel URL ngrok tersebut ke kolom endpoint.

### ⚠️ Catatan penting untuk mode ini:
- DeepSeek-OCR adalah model OCR mentah (bukan LLM instruksi seperti Claude), jadi ia hanya
  mengubah gambar dokumen menjadi teks/tabel markdown. Aplikasi ini memakai parser Python
  (`deepseek_parsers.py`) berbasis pencocokan header tabel untuk mengubah hasil OCR itu
  menjadi data terstruktur.
- Parser dibuat berdasarkan struktur tabel yang **diketahui** dari dokumen GMP (lihat prompt
  di `ocr_engine.py`), tapi **belum dikalibrasi** dengan output nyata DeepSeek-OCR. Jika hasil
  ekstraksi kurang akurat untuk suatu jenis dokumen, jalankan 1 PDF contoh, lihat teks mentah
  yang dikembalikan endpoint `/ocr`, lalu sesuaikan logika pencocokan kolom di
  `deepseek_ocr/parsers.py`.
- Kaggle session GPU gratis punya batas waktu (± 9 jam / run, kuota mingguan terbatas) dan
  URL ngrok berubah setiap kali notebook di-restart — pastikan endpoint di aplikasi diperbarui.

---

## ⚠️ Catatan Penting
1. **Kualitas scan PDF** sangat mempengaruhi akurasi. Pastikan scan jelas dan tidak miring.
2. **Biaya API**: Setiap PDF yang diproses menggunakan Claude API (biaya ~$0.01-0.05 per PDF).
3. **Internet**: Diperlukan koneksi internet untuk memanggil Claude API.
4. **1 PDF = 1 AHU × 1 Semester**: Jangan gabungkan data AHU berbeda dalam 1 PDF.
5. **Chart**: Sistem menggunakan Bar Chart biasa (bukan PivotChart) karena keterbatasan library.
   Jika ingin mengubah ke PivotChart, buka Excel dan buat secara manual dari sheet Table.
6. **xlwings (baru)**: Sistem sekarang mendukung rendering chart via Excel engine (`xlwings`) agar tampilan mendekati PivotChart.
   - Jika Microsoft Excel tidak tersedia / otomatisasi gagal, sistem otomatis fallback ke output `openpyxl` biasa.
   - Untuk menonaktifkan mode `xlwings`, set environment variable: `GMP_USE_XLWINGS=0`

---

## 🔑 Troubleshooting
| Masalah | Solusi |
|---------|--------|
| "Python is not recognized" | Pastikan Python sudah terinstall dan ada di PATH |
| "poppler not found" | Install Poppler sesuai instruksi di atas |
| "API Key error" | Pastikan API key benar dan masih aktif |
| "OCR result incorrect" | Pastikan PDF scan jelas, tidak miring, resolusi cukup |
| "Module not found" | Jalankan: `pip install -r requirements.txt` |

# 📒 AutoLog - Logbook Generator

Personal logbook generator untuk membuat logbook kegiatan selama magang secara otomatis.

## 📝 Deskripsi

AutoLog adalah script Python sederhana yang membantu generate file logbook dalam format Word (.docx) secara otomatis. Cukup masukkan screenshot kegiatan dengan format nama file tertentu, dan script akan membuat logbook lengkap dengan tabel yang rapi, bulan pelaporan otomatis/manual, dan layout A4 yang sudah tersetting.

## ✨ Fitur

- 📅 Generate logbook otomatis berdasarkan bulan saat ini atau bulan pilihan pengguna.
- 🖼️ Import gambar/screenshot kegiatan dari folder dan di-resize otomatis.
- 📄 Output dalam format Word (.docx) dengan layout kertas A4.
- 🔄 Hari otomatis menyesuaikan dengan kalender nyata! (Tidak diulang secara statis).
- 📋 Template logbook yang bisa dikustomisasi.
- ⚡ Smart Launcher (`.bat`) untuk pengguna Windows agar generate logbook bisa dilakukan dengan 1 klik.

---

## 🚀 Cara Penggunaan untuk Pengguna Lain (Clone & Pakai)

Jika Anda men-clone repository ini dan ingin menggunakannya, Anda perlu melakukan beberapa setup dan penyesuaian:

### 1. Setup Environment

```bash
# Clone repository ini
git clone https://github.com/USERNAME_ANDA/AutoLog.git
cd AutoLog

# Buat virtual environment
python -m venv venv

# Aktivasi virtual environment
# Windows (PowerShell):
.\venv\Scripts\Activate.ps1
# Windows (CMD):
venv\Scripts\activate.bat
# Mac / Linux:
source venv/bin/activate

# Install dependencies yang dibutuhkan
python -m pip install --upgrade pip
python -m pip install python-docx docxtpl
```

### 2. Penyesuaian Script (Wajib Diubah!)

Anda harus mengubah beberapa baris kode agar sesuai dengan data Anda:

**A. Pada `generate_logbook.py`**
Buka file `generate_logbook.py` menggunakan text editor kesukaan Anda dan ubah variabel nama pengguna di sekitar **baris 42**:
```python
# Tentukan nama Anda (Bisa diganti sesuai nama pengguna)
NAMA_PENGGUNA = "Nama_Anda"
```

**B. Pada file Launcher `Update_logbook.bat`** (Khusus pengguna Windows)
Jika Anda ingin menggunakan launcher 1-klik, Anda **WAJIB** mengganti path direktori folder AutoLog menuju path Anda sendiri di komputer. Buka `Update_logbook.bat` dengan Notepad dan ubah **baris 16**:
```bat
:: Pindah ke folder script (GANTI PATH INI DENGAN PATH FOLDER ANDA TEMPAT MENYIMPAN AUTOLOG)
cd /d "C:\Users\NamaUser\Documents\AutoLog"
```

### 3. Buat Layout Template Logbook

Sesuaikan desain header perusahaan Anda atau kolom tabel logbook di file `template_logbook.docx`. Anda bisa membuat ulang template mentahnya dengan menjalankan:
```bash
python bikin_template.py
```
*(Pastikan Anda tidak mengubah variable penempat teks seperti `{{ judul_logbook }}` pada template jika Anda tidak merevisinya juga di python).*

### 4. Siapkan Gambar Logging Anda

Buat folder bernama `log_assets/` di dalam root project. Kemudian masukkan gambar (bukti kegiatan) Anda ke dalamnya dengan format nama seperti ini:
```
[Tanggal]_[Deskripsi Kegiatan].png
```

Contoh yang benar:
- `3_Mengerjakan halaman login.png`
- `14_Meeting dengan tim.jpg`
- `25_Review code backend.png`

_Catatan: Tanggalnya harus angka aktual dari tanggal Anda melakukan aktivitas di bulan tersebut._

### 5. Generate Logbook

Gunakan **Smart Launcher `.bat`** (Khusus Windows):
Cukup jalankan file `Update_logbook.bat`. Script akan menanyakan bulan ke-berapa, Anda bisa mengisi `1` untuk Januari, dll., atau langsung pencet Enter untuk menggunakan bulan berjalan. Logbook akan otomatis di-generate dan langsung dibuka di akhir proses.

Atau jalankan secara **Manual via Terminal**:
```bash
python generate_logbook.py [Bulan_Angka]
# Contoh: python generate_logbook.py 2 (untuk generate logbook Februari 2026)
```

Output file yang dihasilkan akan bernama mirip seperti: `Logbook_Februari_Nama_Anda.docx`.

## 📁 Struktur Project

```
AutoLog/
├── bikin_template.py      # Script untuk membuat template statis
├── generate_logbook.py    # Script utama untuk memproses gambar jadi logbook
├── template_logbook.docx  # Template Word yang digunakan sebagai dasaran
├── Update_logbook.bat     # Smart Launcher 1-klik untuk pengguna Windows
├── log_assets/            # (Buat manual) Folder untuk screenshot kegiatan
└── Logbook_*.docx         # Output logbook hasil generate
```

## 📋 Requirements

- Python 3.6+
- `python-docx`
- `docxtpl`

## 👤 Author

**Richie Lagito**

---

_Made with ❤️ untuk mempermudah pembuatan logbook_

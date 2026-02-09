# 📒 AutoLog - Logbook Generator

Personal logbook generator untuk membuat logbook kegiatan selama magang secara otomatis.

## 📝 Deskripsi

AutoLog adalah script Python sederhana yang membantu generate file logbook dalam format Word (.docx) secara otomatis. Cukup masukkan screenshot kegiatan dengan format nama file tertentu, dan script akan membuat logbook lengkap dengan tabel yang rapi.

## ✨ Fitur

- 📅 Generate logbook otomatis berdasarkan bulan saat ini
- 🖼️ Import gambar/screenshot kegiatan dari folder
- 📄 Output dalam format Word (.docx)
- 🔄 Hari otomatis loop Senin-Jumat (hari kerja)
- 📋 Template logbook yang bisa dikustomisasi

## 🚀 Cara Penggunaan

### 1. Setup Environment

```bash
# Buat virtual environment
python3 -m venv venv

# Aktivasi virtual environment
source venv/bin/activate  # Mac/Linux
# atau
venv\Scripts\activate  # Windows

# Install dependencies
pip install python-docx docxtpl
```

### 2. Buat Template

```bash
python3 bikin_template.py
```

Ini akan membuat file `template_logbook.docx`.

### 3. Siapkan Gambar

Buat folder `log_assets/` dan masukkan gambar dengan format nama:

```
[tanggal]_[deskripsi kegiatan].png
```

Contoh:

- `03_Mengerjakan halaman login.png`
- `04_Meeting dengan tim.jpg`
- `05_Review code backend.png`

### 4. Generate Logbook

```bash
python3 generate_logbook.py
```

Output: `Logbook_[Bulan]_Richie.docx`

## 📁 Struktur Project

```
AutoLog/
├── bikin_template.py      # Script untuk membuat template
├── generate_logbook.py    # Script utama untuk generate logbook
├── template_logbook.docx  # Template yang di-generate
├── log_assets/            # Folder untuk screenshot kegiatan
└── Logbook_*.docx         # Output logbook
```

## 📋 Requirements

- Python 3.6+
- python-docx
- docxtpl

## 👤 Author

**Richie Lagito**

---

_Made with ❤️ untuk mempermudah pembuatan logbook magang_

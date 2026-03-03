import os
import sys
import calendar
from datetime import datetime
from docxtpl import DocxTemplate
from docx import Document 
from docx.shared import Mm, Cm

def generate_logbook():
    # --- 1. KONFIGURASI ---
    template_path = "template_logbook.docx"
    image_folder = "./log_assets"
    
    # Mapping Nama Bulan (Inggris -> Indonesia)
    bulan_indo = {
        1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April', 5: 'Mei', 6: 'Juni',
        7: 'Juli', 8: 'Agustus', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
    }
    
    now = datetime.now()
    
    # Ambil bulan dari parameter command line (lewat .bat)
    if len(sys.argv) > 1 and sys.argv[1].isdigit():
        bulan_target = int(sys.argv[1])
        if bulan_target < 1 or bulan_target > 12:
            print(f"Bulan {bulan_target} tidak valid! Menggunakan bulan ini ({now.month}).")
            bulan_target = now.month
    else:
        bulan_target = now.month # Default ke bulan saat ini

    # Fix tahun ke 2026
    tahun_target = 2026
    
    nama_bulan_ini = bulan_indo[bulan_target]
    tahun_ini = tahun_target
    
    # Tanggal dan bulan tanda tangan menggunakan waktu saat ini digenerate
    tanggal_tanda_tangan = now.day
    nama_bulan_ttd = bulan_indo[now.month]
    tahun_ttd = now.year
    
    # Tentukan nama Anda (Bisa diganti sesuai nama pengguna)
    NAMA_PENGGUNA = "Nikolaus"
    
    # Output path dinamis berdasarkan bulan yang dipilih
    output_path = f"Logbook_{nama_bulan_ini}_{NAMA_PENGGUNA}.docx"
    
    # Cek kelengkapan
    if not os.path.exists(template_path):
        print(f"ERROR: File '{template_path}' tidak ditemukan. Jalankan script pembuat template dulu!")
        return
    if not os.path.exists(image_folder):
        print(f"ERROR: Folder '{image_folder}' tidak ditemukan. Buat folder dan isi gambar dulu!")
        return

    # --- 2. RENDER HEADER/FOOTER DULU (DOCXTPL) ---
    doc_tpl = DocxTemplate(template_path)
    
    # Context untuk variabel Header/Footer
    # judul_logbook menggunakan bulan yang diinput
    # tanggal_ttd menggunakan waktu saat file ini digenerate
    context = {
        'judul_logbook': f"{nama_bulan_ini} {tahun_ini}", 
        'tanggal_ttd': f"{tanggal_tanda_tangan} {nama_bulan_ttd} {tahun_ttd}" 
    }
    
    print("Merenender Header/Footer...")
    doc_tpl.render(context)
    temp_path = "temp_logbook.docx"
    doc_tpl.save(temp_path)
    
    # --- 3. ISI TABEL MANUAL (PYTHON-DOCX) ---
    print(f"Mengisi tabel dengan gambar untuk bulan {nama_bulan_ini} {tahun_ini}...")
    doc = Document(temp_path)
    
    # Set page size ke A4 (210mm x 297mm)
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
    
    # Fungsi Helper untuk Styling Text (Font & Alignment)
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    
    def fill_cell(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False):
        cell.text = text
        p = cell.paragraphs[0]
        p.alignment = align
        run = p.runs[0]
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = bold

    # Ambil tabel pertama
    table = doc.tables[0]
    
    # Mapping Nama Hari
    nama_hari_indo = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']

    # Ambil data gambar
    files = sorted([f for f in os.listdir(image_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))])
    print(f"Ditemukan {len(files)} file gambar.")

    for index, filename in enumerate(files, start=1):
        try:
            parts = filename.split('_', 1) 
            if len(parts) < 2:
                continue
                
            tgl_str = parts[0]
            desc_raw = parts[1]
            kegiatan = desc_raw.rsplit('.', 1)[0]
            
            # Cari hari berdasarkan tanggal aktual
            tgl_int = int(tgl_str)
            try:
                tgl_obj = datetime(tahun_ini, bulan_target, tgl_int)
                nama_hari = nama_hari_indo[tgl_obj.weekday()]
            except ValueError:
                nama_hari = "Hari Tidak Valid"
                
            # Full_date di tabel menggunakan bulan yang diset (bulan_target)
            full_date = f"{nama_hari}, {tgl_int} {nama_bulan_ini} {tahun_ini}"
            
            # Tambah Row Baru
            row = table.add_row()
            
            # Isi Cell dengan Styling
            fill_cell(row.cells[0], str(index), align=WD_ALIGN_PARAGRAPH.CENTER)
            fill_cell(row.cells[1], full_date, align=WD_ALIGN_PARAGRAPH.CENTER)
            fill_cell(row.cells[2], kegiatan, align=WD_ALIGN_PARAGRAPH.LEFT)
            
            # Col 3: Gambar (Center)
            cell_img = row.cells[3]
            p_img = cell_img.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.clear()
            run_img = p_img.add_run()
            
            img_path = os.path.join(image_folder, filename)
            run_img.add_picture(img_path, width=Mm(60)) 
            
            print(f"  [OK] No {index}: {kegiatan}")

        except ValueError as e:
            print(f"Warning: Error pada file '{filename}': {e}")

    doc.save(output_path)
    
    # Hapus file temp
    if os.path.exists(temp_path):
        os.remove(temp_path)
        
    print(f"\nSUKSES! File tersimpan di: {output_path}")

if __name__ == "__main__":
    generate_logbook()

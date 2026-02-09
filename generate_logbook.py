import os
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
    nama_bulan_ini = bulan_indo[now.month]
    tahun_ini = now.year
    
    # Output path dinamis berdasarkan bulan
    output_path = f"Logbook_{nama_bulan_ini}_Richie.docx"
    
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
    context = {
        'judul_logbook': f"{nama_bulan_ini} {tahun_ini}", # Otomatis!
        'tanggal_ttd': f"{now.day} {nama_bulan_ini} {tahun_ini}" 
    }
    
    print("Merenender Header/Footer...")
    doc_tpl.render(context)
    temp_path = "temp_logbook.docx"
    doc_tpl.save(temp_path)
    
    # --- 3. ISI TABEL MANUAL (PYTHON-DOCX) ---
    print("Mengisi tabel dengan gambar...")
    doc = Document(temp_path)
    
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
    
    # List hari kerja (Senin - Jumat) untuk looping
    hari_kerja = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']

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
            
            # Loop hari kerja: Senin, Selasa, Rabu, Kamis, Jumat, Senin, ...
            nama_hari = hari_kerja[(index - 1) % 5]
            full_date = f"{nama_hari}, {int(tgl_str)} {nama_bulan_ini} {tahun_ini}"
            
            # Tambah Row Baru
            row = table.add_row()
            
            # Isi Cell dengan Styling
            # Col 0: No (Center)
            fill_cell(row.cells[0], str(index), align=WD_ALIGN_PARAGRAPH.CENTER)
            
            # Col 1: Tanggal (Center/Left? Kita set Center biar rapi)
            fill_cell(row.cells[1], full_date, align=WD_ALIGN_PARAGRAPH.CENTER)
            
            # Col 2: Kegiatan (Left, Justify)
            fill_cell(row.cells[2], kegiatan, align=WD_ALIGN_PARAGRAPH.LEFT)
            
            # Col 3: Gambar (Center)
            cell_img = row.cells[3]
            p_img = cell_img.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER # Center Gambar
            p_img.clear()
            run_img = p_img.add_run()
            
            img_path = os.path.join(image_folder, filename)
            # EDIT UKURAN GAMBAR DISINI (width=Mm(45))
            run_img.add_picture(img_path, width=Mm(60)) 
            
            print(f"  [OK] No {index}: {kegiatan}")

        except ValueError as e:
            print(f"Warning: Error pada file '{filename}': {e}")
            
    # Set Width (Penting diulang agar tabel tidak 'layu')
    widths = [Cm(1), Cm(4), Cm(6), Cm(5)]
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

    doc.save(output_path)
    
    # Hapus file temp
    if os.path.exists(temp_path):
        os.remove(temp_path)
        
    print(f"\nSUKSES! File tersimpan di: {output_path}")

if __name__ == "__main__":
    generate_logbook()
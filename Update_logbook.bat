@echo off
title Auto Logbook PUSDATIN - Smart Launcher
color 0b

echo ===================================================
echo             AUTO LOGBOOK GENERATOR 2026
echo ===================================================
echo.
set /p bulan_input="Masukkan Angka Bulan untuk Logbook (1-12) [Enter untuk bulan ini]: "

echo.
echo Menuju folder OneDrive PUSDATIN...

:: Pindah ke folder script di OneDrive (GANTI PATH INI DENGAN PATH FOLDER ANDA)
:: Contoh: cd /d "C:\Users\NamaUser\Documents\AutoLog"
cd /d "C:\Users\PUSDATIN 01\OneDrive\OneDrive - Universitas Tarumanagara\PUSDATIN\Auto Logbook\AutoLog"

echo Menjalankan Auto Logbook...
:: Sisipkan input bulan target ke dalam argument Python
powershell -ExecutionPolicy Bypass -Command ".\venv\Scripts\Activate.ps1; python generate_logbook.py %bulan_input%"

echo.
echo Mencari dan membuka file logbook terbaru...

:: Buka file Logbook terbaru berdasarkan waktu modifikasi terakhir
powershell -Command "$latest = Get-ChildItem -Filter 'Logbook_*.docx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($latest) { start $latest.FullName } else { echo 'File tidak ditemukan!' }"

echo.
echo Proses Selesai! Menutup dalam...
timeout /t 3
exit

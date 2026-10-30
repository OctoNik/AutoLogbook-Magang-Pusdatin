@echo off
title Auto Logbook PUSDATIN - Smart Launcher
color 0b

echo ===================================================
echo             AUTO LOGBOOK GENERATOR 2026
echo ===================================================
echo.

:: 1. Cek apakah Python terinstall
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python tidak ditemukan di laptop ini!
    echo Silakan install Python terlebih dahulu dan pastikan mencentang "Add Python to PATH".
    echo.
    pause
    exit /b
)

set /p bulan_input="Masukkan Angka Bulan untuk Logbook (1-12) [Enter untuk bulan ini]: "

echo.
echo Menuju folder script...

:: Pindah ke folder script saat ini secara dinamis (menggunakan %~dp0)
cd /d "%~dp0"

:: 2. Cek apakah library python-docx, docxtpl, dan holidays sudah terinstall di sistem
echo Memeriksa library pendukung...
python -c "import docxtpl, docx, holidays" >nul 2>nul
if %errorlevel% neq 0 (
    echo Library belum lengkap. Menginstal python-docx, docxtpl, dan holidays ke Python sistem...
    python -m pip install python-docx docxtpl holidays
    if %errorlevel% neq 0 (
        echo [ERROR] Gagal menginstal library! Pastikan laptop terhubung ke internet.
        pause
        exit /b
    )
)

echo Menjalankan Auto Logbook...
:: Jalankan langsung dengan Python sistem
python generate_logbook.py %bulan_input%

echo.
echo Mencari dan membuka file logbook terbaru...

:: Buka file Logbook terbaru berdasarkan waktu modifikasi terakhir
powershell -Command "$latest = Get-ChildItem -Filter 'Logbook_*.docx' | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($latest) { start $latest.FullName } else { echo 'File tidak ditemukan!' }"

echo.
echo Proses Selesai! Menutup dalam...
timeout /t 3
exit

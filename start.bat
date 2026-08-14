@echo off
setlocal enabledelayedexpansion

echo ============================================
echo   Excel Toolkit - Menjalankan Aplikasi
echo ============================================
echo.

if not exist venv (
    echo [ERROR] Virtual environment belum ada.
    echo Jalankan install_dependencies.bat terlebih dahulu.
    pause
    exit /b 1
)

call venv\Scripts\activate.bat

echo Alamat IP yang dapat diakses dari jaringan yang sama:
echo.
set "found="
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /R /C:"IPv4"') do (
    set "ip=%%a"
    set "ip=!ip:~1!"           REM hilangkan spasi di depan
    if not "!ip!"=="127.0.0.1" (
        echo    http://!ip!:5000
        set "found=1"
    )
)
if not defined found (
    echo    (Tidak ada alamat IP selain localhost. Pastikan terhubung ke jaringan.)
)
echo.
echo Untuk akses di komputer ini (localhost):
echo    http://localhost:5000
echo.
echo (Bagikan salah satu alamat IP di atas ke pengguna lain di WiFi yang sama.)
echo.

start "" http://localhost:5000

echo Menjalankan server Flask...
echo Tekan CTRL+C untuk menghentikan server.
echo.

python app.py

pause
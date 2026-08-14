@echo off
setlocal

echo ============================================
echo   Excel Toolkit - Instalasi Dependensi
echo ============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan di PATH.
    echo Silakan install Python 3.10+ dari https://www.python.org/downloads/
    echo Pastikan mencentang "Add Python to PATH" saat instalasi.
    pause
    exit /b 1
)

echo [1/4] Python ditemukan:
python --version
echo.

if not exist venv (
    echo [2/4] Membuat virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo [ERROR] Gagal membuat virtual environment.
        pause
        exit /b 1
    )
) else (
    echo [2/4] Virtual environment sudah ada, dilewati.
)
echo.

echo [3/4] Mengaktifkan virtual environment dan upgrade pip...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip

echo.
echo [4/4] Menginstall dependensi dari requirements.txt...
pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo [ERROR] Instalasi dependensi gagal. Periksa pesan error di atas.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Instalasi selesai! Jalankan run_app.bat
echo   untuk memulai aplikasi.
echo ============================================
pause
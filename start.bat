@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo  Klasifikasi Keluhan Transportasi - START
echo ============================================

if not exist venv (
    echo [1/4] Membuat virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo GAGAL: pastikan Python sudah terinstall dan ada di PATH.
        pause
        exit /b 1
    )
) else (
    echo [1/4] Virtual environment sudah ada, lewati pembuatan.
)

echo [2/4] Mengaktifkan virtual environment...
call venv\Scripts\activate.bat

echo [3/4] Menginstall/memeriksa library...
pip install -q -r requirements.txt
if errorlevel 1 (
    echo GAGAL: instalasi library bermasalah, cek koneksi internet atau requirements.txt.
    pause
    exit /b 1
)

echo [4/4] Menjalankan server di http://localhost:8000 ...
start "KeluhanTransportasiServer" cmd /k venv\Scripts\python -m uvicorn aplikasi_web.main:app --port 8000

timeout /t 3 /nobreak >nul
start "" http://localhost:8000

echo.
echo Server berjalan di jendela terpisah berjudul "KeluhanTransportasiServer".
echo Jalankan stop.bat untuk menghentikan server.
endlocal

@echo off
setlocal
echo ============================================
echo  Klasifikasi Keluhan Transportasi - STOP
echo ============================================

echo Menghentikan proses yang memakai port 8000...
for /f "tokens=5" %%p in ('netstat -aon ^| findstr :8000 ^| findstr LISTENING') do (
    taskkill /PID %%p /F >nul 2>&1
)

echo Menutup jendela server (jika masih terbuka)...
taskkill /FI "WINDOWTITLE eq KeluhanTransportasiServer*" /T /F >nul 2>&1

echo Selesai. Server sudah dihentikan.
pause
endlocal

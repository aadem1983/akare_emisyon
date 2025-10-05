@echo off
echo === EMISYON SAHA UYGULAMASI (GELISTIRME MODU) ===
echo.

REM Proje dizinine git
cd /d "%~dp0"
echo Proje dizini: %CD%

REM Python kontrol
python --version
if %errorlevel% neq 0 (
    echo HATA: Python bulunamadı!
    pause
    exit /b 1
)

REM Flask kontrol
python -c "import flask" 2>nul
if %errorlevel% neq 0 (
    echo Flask yükleniyor...
    pip install -r requirements.txt
)

REM Environment variables ayarla
set PORT=5001
set FLASK_ENV=development
set FLASK_DEBUG=1
set FLASK_APP=app.py

echo.
echo === UYGULAMA BAŞLATILIYOR ===
echo URL: http://localhost:5001
echo Debug Mode: AÇIK
echo Auto-Reload: AÇIK
echo Durdurmak için Ctrl+C tuşlayın
echo.

REM Flask uygulamasını başlat
python app.py

echo.
echo Uygulama durduruldu.
pause





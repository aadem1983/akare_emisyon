@echo off
echo === EMISYON SAHA UYGULAMASI BAŞLATILIYOR ===
echo.

REM Proje dizinine git
cd /d "C:\Users\FOCUSGC\Desktop\AKARE-YAZILIM\EMISYON SAHA\CURSOR-EMISYON\CURSAR-EMISYON-1"
echo Proje dizinine gidiliyor...

REM Python kontrol
python --version
if %errorlevel% neq 0 (
    echo HATA: Python bulunamadı!
    pause
    exit
)

REM Flask kontrol ve yükleme
python -c "import flask" 2>nul
if %errorlevel% neq 0 (
    echo Flask yükleniyor...
    pip install -r requirements.txt
)

REM Uygulamayı başlat
echo.
echo === UYGULAMA BAŞLATILIYOR ===
echo Tarayıcınızda http://localhost:5001 adresini açın
echo Durdurmak için Ctrl+C tuşlayın
echo.

set PORT=5001
set FLASK_ENV=development
set FLASK_DEBUG=1
python app.py

echo.
echo Uygulama durduruldu.
pause

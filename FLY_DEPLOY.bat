@echo off
echo ========================================
echo    FLY.IO DEPLOY SCRIPT
echo ========================================
echo.

cd /d "C:\Users\FOCUSGC\Desktop\AKARE-YAZILIM\EMISYON SAHA\CURSOR-EMISYON\CURSAR-EMISYON-1"

echo [1/3] GitHub'a push yapiliyor...
git add .
git commit -m "Deploy: %date% %time%"
git push origin main

if %errorlevel% neq 0 (
    echo HATA: GitHub push basarisiz!
    pause
    exit /b 1
)

echo [2/3] Fly.io'ya deploy ediliyor...
flyctl deploy

if %errorlevel% neq 0 (
    echo HATA: Fly.io deploy basarisiz!
    pause
    exit /b 1
)

echo [3/3] Deploy durumu kontrol ediliyor...
flyctl status

echo.
echo ========================================
echo    DEPLOY TAMAMLANDI!
echo ========================================
echo.
echo Uygulamaniz: https://akare-emisyon.fly.dev/
echo.
pause

# Emisyon Saha Uygulaması Başlatma Script'i

Write-Host "=== EMISYON SAHA UYGULAMASI BAŞLATILIYOR ===" -ForegroundColor Green
Write-Host ""

# Proje dizinine git
$projectPath = "C:\Users\FOCUSGC\Desktop\AKARE-YAZILIM\EMISYON SAHA\CURSOR-EMISYON\CURSAR-EMISYON-1"
Write-Host "Proje dizinine gidiliyor: $projectPath" -ForegroundColor Yellow

if (Test-Path $projectPath) {
    Set-Location $projectPath
    Write-Host "Proje dizinine başarıyla gidildi." -ForegroundColor Green
} else {
    Write-Host "HATA: Proje dizini bulunamadı!" -ForegroundColor Red
    pause
    exit
}

# Python kontrol
Write-Host "Python kontrol ediliyor..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python sürümü: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "HATA: Python bulunamadı!" -ForegroundColor Red
    pause
    exit
}

# Flask kontrol
Write-Host "Flask kontrol ediliyor..." -ForegroundColor Yellow
try {
    python -c "import flask" 2>$null
    Write-Host "Flask yüklü" -ForegroundColor Green
} catch {
    Write-Host "Flask yükleniyor..." -ForegroundColor Yellow
    pip install -r requirements.txt
}

# Uygulamayı başlat
Write-Host ""
Write-Host "=== UYGULAMA BAŞLATILIYOR ===" -ForegroundColor Green
Write-Host "Tarayıcınızda http://localhost:5001 adresini açın" -ForegroundColor Cyan
Write-Host "Durdurmak için Ctrl+C tuşlayın" -ForegroundColor Yellow
Write-Host ""

$env:PORT=5001
$env:FLASK_ENV="development"
$env:FLASK_DEBUG="1"
python app.py

Write-Host ""
Write-Host "Uygulama durduruldu." -ForegroundColor Yellow
pause
# Emisyon Saha Uygulaması Başlatma Script'i
# Bu script Flask uygulamasını localhost:5001'de başlatır

Write-Host "=== EMISYON SAHA UYGULAMASI BAŞLATILIYOR ===" -ForegroundColor Green
Write-Host ""

# Proje dizinine git
$projectPath = "C:\Users\FOCUSGC\Desktop\AKARE-YAZILIM\EMISYON SAHA\CURSOR-EMISYON\CURSAR-EMISYON-1"
Write-Host "Proje dizinine gidiliyor: $projectPath" -ForegroundColor Yellow

# Dizinin var olup olmadığını kontrol et
if (Test-Path $projectPath) {
    Set-Location $projectPath
    Write-Host "Proje dizinine başarıyla gidildi." -ForegroundColor Green
} else {
    Write-Host "HATA: Proje dizini bulunamadı!" -ForegroundColor Red
    Write-Host "Lütfen dizin yolunu kontrol edin." -ForegroundColor Red
    pause
    exit
}

# Python'un yüklü olup olmadığını kontrol et
Write-Host "Python kontrol ediliyor..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python sürümü: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "HATA: Python bulunamadı!" -ForegroundColor Red
    Write-Host "Lütfen Python'u yükleyin." -ForegroundColor Red
    pause
    exit
}

# Gerekli paketlerin yüklü olup olmadığını kontrol et
Write-Host "Gerekli paketler kontrol ediliyor..." -ForegroundColor Yellow
try {
    python -c "import flask" 2>$null
    Write-Host "Flask yüklü ✓" -ForegroundColor Green
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

# Flask uygulamasını başlat
python app.py

# Uygulama durduğunda
Write-Host ""
Write-Host "Uygulama durduruldu." -ForegroundColor Yellow
pause

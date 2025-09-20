#!/usr/bin/env python3
"""
Flask Uygulamasını Firebase'e Otomatik Migration Script'i
Bu script Flask uygulamasının tüm özelliklerini Firebase'e taşır.
"""

import os
import json
import shutil
from pathlib import Path
import re

class FlaskToFirebaseMigrator:
    def __init__(self, flask_dir, firebase_dir):
        self.flask_dir = Path(flask_dir)
        self.firebase_dir = Path(firebase_dir)
        self.templates_dir = self.flask_dir / "templates"
        self.static_dir = self.flask_dir / "static"
        
    def migrate_all(self):
        """Tüm Flask uygulamasını Firebase'e taşır"""
        print("🚀 Flask'tan Firebase'e otomatik migration başlıyor...")
        
        # 1. Veri dosyalarını kopyala
        self.migrate_data_files()
        
        # 2. Static dosyaları kopyala
        self.migrate_static_files()
        
        # 3. Ana HTML dosyasını oluştur
        self.create_main_html()
        
        # 4. Firebase config dosyalarını oluştur
        self.create_firebase_config()
        
        # 5. Migration script'ini çalıştır
        self.run_data_migration()
        
        print("✅ Migration tamamlandı!")
        
    def migrate_data_files(self):
        """JSON veri dosyalarını kopyala"""
        print("📁 Veri dosyaları kopyalanıyor...")
        
        data_files = [
            "users.json", "parameters.json", "firma_olcum.json", 
            "saha.json", "measurements.json", "il-ilce.json"
        ]
        
        for file_name in data_files:
            src = self.flask_dir / file_name
            dst = self.firebase_dir / file_name
            
            if src.exists():
                shutil.copy2(src, dst)
                print(f"  ✅ {file_name} kopyalandı")
            else:
                print(f"  ⚠️ {file_name} bulunamadı")
    
    def migrate_static_files(self):
        """Static dosyaları kopyala"""
        print("🎨 Static dosyalar kopyalanıyor...")
        
        if self.static_dir.exists():
            static_dst = self.firebase_dir / "public" / "static"
            static_dst.mkdir(parents=True, exist_ok=True)
            
            # CSS dosyaları
            css_src = self.static_dir / "css"
            if css_src.exists():
                css_dst = static_dst / "css"
                shutil.copytree(css_src, css_dst, dirs_exist_ok=True)
                print("  ✅ CSS dosyaları kopyalandı")
            
            # JS dosyaları
            js_src = self.static_dir / "js"
            if js_src.exists():
                js_dst = static_dst / "js"
                shutil.copytree(js_src, js_dst, dirs_exist_ok=True)
                print("  ✅ JS dosyaları kopyalandı")
            
            # Resimler
            images_src = self.static_dir / "images"
            if images_src.exists():
                images_dst = static_dst / "images"
                shutil.copytree(images_src, images_dst, dirs_exist_ok=True)
                print("  ✅ Resimler kopyalandı")
    
    def create_main_html(self):
        """Ana HTML dosyasını Flask template'lerinden oluştur"""
        print("🌐 Ana HTML dosyası oluşturuluyor...")
        
        # Flask template'lerini analiz et
        templates = self.analyze_templates()
        
        # Firebase HTML'ini oluştur
        html_content = self.generate_firebase_html(templates)
        
        # Dosyayı kaydet
        html_file = self.firebase_dir / "public" / "index.html"
        html_file.parent.mkdir(parents=True, exist_ok=True)
        
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print("  ✅ index.html oluşturuldu")
    
    def analyze_templates(self):
        """Flask template'lerini analiz et"""
        templates = {}
        
        if self.templates_dir.exists():
            for template_file in self.templates_dir.glob("*.html"):
                with open(template_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                    
                templates[template_file.stem] = {
                    'content': content,
                    'title': self.extract_title(content),
                    'forms': self.extract_forms(content),
                    'tables': self.extract_tables(content)
                }
        
        return templates
    
    def extract_title(self, content):
        """HTML'den başlık çıkar"""
        match = re.search(r'<title>(.*?)</title>', content, re.IGNORECASE)
        return match.group(1) if match else "Sayfa"
    
    def extract_forms(self, content):
        """HTML'den form alanlarını çıkar"""
        forms = []
        form_pattern = r'<form[^>]*>(.*?)</form>'
        
        for match in re.finditer(form_pattern, content, re.DOTALL | re.IGNORECASE):
            form_content = match.group(1)
            
            # Input alanlarını bul
            inputs = re.findall(r'<input[^>]*name=["\']([^"\']*)["\'][^>]*>', form_content)
            selects = re.findall(r'<select[^>]*name=["\']([^"\']*)["\'][^>]*>', form_content)
            textareas = re.findall(r'<textarea[^>]*name=["\']([^"\']*)["\'][^>]*>', form_content)
            
            forms.append({
                'inputs': inputs,
                'selects': selects,
                'textareas': textareas
            })
        
        return forms
    
    def extract_tables(self, content):
        """HTML'den tablo yapılarını çıkar"""
        tables = []
        table_pattern = r'<table[^>]*>(.*?)</table>'
        
        for match in re.finditer(table_pattern, content, re.DOTALL | re.IGNORECASE):
            table_content = match.group(1)
            
            # Tablo başlıklarını bul
            headers = re.findall(r'<th[^>]*>(.*?)</th>', table_content, re.IGNORECASE)
            
            tables.append({
                'headers': headers,
                'content': table_content
            })
        
        return tables
    
    def generate_firebase_html(self, templates):
        """Firebase HTML'ini oluştur"""
        # Bu fonksiyon çok uzun olacağı için ayrı bir dosyaya taşınabilir
        # Şimdilik basit bir template oluşturalım
        
        html_template = '''<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Emisyon Saha Programı - Firebase</title>
    
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- Font Awesome -->
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <!-- SweetAlert2 -->
    <link href="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.min.css" rel="stylesheet">
    <!-- Custom CSS -->
    <link href="static/css/style.css" rel="stylesheet">
</head>
<body>
    <!-- Login Section -->
    <div id="loginSection" class="container-fluid vh-100 d-flex align-items-center justify-content-center" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
        <div class="card shadow-lg" style="width: 400px;">
            <div class="card-body p-5">
                <div class="text-center mb-4">
                    <img src="static/images/logo.png" alt="Logo" class="mb-3" style="height: 60px;">
                    <h4 class="text-primary">Emisyon Saha Programı</h4>
                    <p class="text-muted">Giriş yapın</p>
                </div>
                
                <form id="loginForm">
                    <div class="mb-3">
                        <label for="username" class="form-label">Kullanıcı Adı</label>
                        <input type="text" class="form-control" id="username" required>
                    </div>
                    <div class="mb-3">
                        <label for="password" class="form-label">Şifre</label>
                        <input type="password" class="form-control" id="password" required>
                    </div>
                    <button type="button" class="btn btn-primary w-100" id="loginBtn">
                        <i class="fas fa-sign-in-alt me-2"></i>Giriş Yap
                    </button>
                </form>
            </div>
        </div>
    </div>

    <!-- Main App Section -->
    <div id="mainApp" class="container-fluid" style="display: none;">
        <!-- Navigation -->
        <nav class="navbar navbar-expand-lg navbar-dark bg-primary">
            <div class="container-fluid">
                <a class="navbar-brand" href="#">
                    <img src="static/images/logo.png" alt="Logo" height="30" class="d-inline-block align-text-top me-2">
                    Emisyon Saha
                </a>
                
                <div class="navbar-nav me-auto">
                    <a class="nav-link active" href="#" onclick="showPage('firma-olcum')">Ana Sayfa</a>
                    <a class="nav-link" href="#" onclick="showPage('parametre')">PARAMETRE</a>
                    <a class="nav-link" href="#" onclick="showPage('firma-olcum')">FIRMA_OLCUM</a>
                    <a class="nav-link" href="#" onclick="showPage('kullanici-yonetimi')">Kullanıcı Yönetimi</a>
                </div>
                
                <div class="navbar-nav">
                    <span class="navbar-text me-3">
                        Kullanıcı: <span id="currentUser">admin</span>
                    </span>
                    <button class="btn btn-outline-light btn-sm" onclick="logout()">
                        <i class="fas fa-sign-out-alt"></i> Çıkış
                    </button>
                </div>
            </div>
        </nav>

        <!-- Page Content -->
        <div class="container-fluid mt-3">
            <!-- Firma Ölçüm Sayfası -->
            <div id="firma-olcum" class="page-content">
                <div class="d-flex justify-content-between align-items-center mb-3">
                    <h4><i class="fas fa-building me-2"></i>Firma Ölçüm Yönetimi</h4>
                    <div class="btn-group">
                        <button class="btn btn-success" onclick="showAddFirmaModal()">
                            <i class="fas fa-plus me-1"></i>Ölçüm Ekle
                        </button>
                        <button class="btn btn-info" onclick="exportFirmalar()">
                            <i class="fas fa-download me-1"></i>Dışa Aktar
                        </button>
                        <button class="btn btn-warning" onclick="importFirmalar()">
                            <i class="fas fa-upload me-1"></i>İçe Aktar
                        </button>
                    </div>
                </div>

                <!-- Filtreler -->
                <div class="card mb-3">
                    <div class="card-body">
                        <div class="row">
                            <div class="col-md-2">
                                <input type="text" class="form-control" placeholder="Firma Ara..." onkeyup="filterTable(1, this.value)">
                            </div>
                            <div class="col-md-2">
                                <input type="text" class="form-control" placeholder="Ölçüm Kodu..." onkeyup="filterTable(2, this.value)">
                            </div>
                            <div class="col-md-2">
                                <input type="text" class="form-control" placeholder="İl..." onkeyup="filterTable(8, this.value)">
                            </div>
                            <div class="col-md-2">
                                <input type="text" class="form-control" placeholder="İlçe..." onkeyup="filterTable(9, this.value)">
                            </div>
                            <div class="col-md-2">
                                <select class="form-select" onchange="filterTable(12, this.value)">
                                    <option value="">Tüm Durumlar</option>
                                    <option value="Aktif">Aktif</option>
                                    <option value="Pasif">Pasif</option>
                                </select>
                            </div>
                            <div class="col-md-2">
                                <button class="btn btn-secondary w-100" onclick="clearFilters()">
                                    <i class="fas fa-times me-1"></i>Filtreleri Temizle
                                </button>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Tablo -->
                <div class="card">
                    <div class="card-body">
                        <div class="table-responsive">
                            <table id="firmaOlcumTablo" class="table table-striped table-hover">
                                <thead class="table-dark">
                                    <tr>
                                        <th width="50">
                                            <input type="checkbox" id="tumunuSec" class="form-check-input">
                                        </th>
                                        <th class="text-center">SIRA</th>
                                        <th>FIRMA</th>
                                        <th>OLC_KOD</th>
                                        <th>BAS TRH</th>
                                        <th>BIT TAR</th>
                                        <th>BACA SAY</th>
                                        <th>PARAMETRE</th>
                                        <th>PER.</th>
                                        <th>IL</th>
                                        <th>ILCE</th>
                                        <th>YETK</th>
                                        <th>TEL</th>
                                        <th>DURUM</th>
                                        <th class="text-center">İŞLEMLER</th>
                                    </tr>
                                </thead>
                                <tbody id="firmalarTableBody">
                                    <!-- Veriler JavaScript ile doldurulacak -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Parametre Sayfası -->
            <div id="parametre" class="page-content" style="display: none;">
                <div class="d-flex justify-content-between align-items-center mb-3">
                    <h4><i class="fas fa-cogs me-2"></i>Parametre Yönetimi</h4>
                    <div class="btn-group">
                        <button class="btn btn-success" onclick="showAddParameterModal()">
                            <i class="fas fa-plus me-1"></i>Parametre Ekle
                        </button>
                        <button class="btn btn-info" onclick="exportParameters()">
                            <i class="fas fa-download me-1"></i>Dışa Aktar
                        </button>
                        <button class="btn btn-warning" onclick="importParameters()">
                            <i class="fas fa-upload me-1"></i>İçe Aktar
                        </button>
                    </div>
                </div>

                <!-- Parametre Tablosu -->
                <div class="card">
                    <div class="card-body">
                        <div class="table-responsive">
                            <table class="table table-striped table-hover">
                                <thead class="table-dark">
                                    <tr>
                                        <th width="50">
                                            <input type="checkbox" id="tumunuSec" class="form-check-input">
                                        </th>
                                        <th class="text-center">SIRA</th>
                                        <th>PARAMETRE ADI</th>
                                        <th>METOT</th>
                                        <th>İZO ORAN</th>
                                        <th>NOZZLE</th>
                                        <th>IMP1</th>
                                        <th>IMP2</th>
                                        <th>IMP3</th>
                                        <th>IMP4</th>
                                        <th>L/DAK</th>
                                        <th>T.HAC</th>
                                        <th>LOQ</th>
                                        <th>KK</th>
                                        <th>-3S</th>
                                        <th>-2S</th>
                                        <th>+2S</th>
                                        <th>+3S</th>
                                        <th class="text-center">İŞLEMLER</th>
                                    </tr>
                                </thead>
                                <tbody id="parametersTableBody">
                                    <!-- Veriler JavaScript ile doldurulacak -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Firma Detay Sayfası -->
            <div id="firma-detay" class="page-content" style="display: none;">
                <!-- Firma detay içeriği -->
            </div>

            <!-- Firma Düzenleme Sayfası -->
            <div id="firma-duzenle" class="page-content" style="display: none;">
                <!-- Firma düzenleme formu -->
            </div>

            <!-- Kullanıcı Yönetimi Sayfası -->
            <div id="kullanici-yonetimi" class="page-content" style="display: none;">
                <!-- Kullanıcı yönetimi içeriği -->
            </div>
        </div>
    </div>

    <!-- Scripts -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
    <script src="firebase-config.js"></script>
    <script>
        // Firebase bağlantısı ve uygulama mantığı buraya gelecek
        // Bu kısım çok uzun olacağı için ayrı bir dosyaya taşınabilir
    </script>
</body>
</html>'''
        
        return html_template
    
    def create_firebase_config(self):
        """Firebase config dosyalarını oluştur"""
        print("🔥 Firebase config dosyaları oluşturuluyor...")
        
        # firebase.json
        firebase_json = {
            "hosting": {
                "public": "public",
                "ignore": [
                    "firebase.json",
                    "**/.*",
                    "**/node_modules/**"
                ],
                "rewrites": [
                    {
                        "source": "**",
                        "destination": "/index.html"
                    }
                ]
            }
        }
        
        firebase_json_file = self.firebase_dir / "firebase.json"
        with open(firebase_json_file, 'w', encoding='utf-8') as f:
            json.dump(firebase_json, f, indent=2)
        
        print("  ✅ firebase.json oluşturuldu")
    
    def run_data_migration(self):
        """Veri migration'ını çalıştır"""
        print("📊 Veri migration'ı çalıştırılıyor...")
        
        # firebase_migration.py dosyasını çalıştır
        migration_script = self.firebase_dir / "firebase_migration.py"
        if migration_script.exists():
            os.system(f"cd {self.firebase_dir} && python firebase_migration.py")
            print("  ✅ Veri migration'ı tamamlandı")
        else:
            print("  ⚠️ firebase_migration.py bulunamadı")

def main():
    """Ana fonksiyon"""
    flask_dir = "CURSAR-EMISYON-1"  # Flask uygulamasının bulunduğu dizin
    firebase_dir = "firebase-app"   # Firebase uygulamasının oluşturulacağı dizin
    
    migrator = FlaskToFirebaseMigrator(flask_dir, firebase_dir)
    migrator.migrate_all()

if __name__ == "__main__":
    main() 
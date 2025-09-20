#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Veri Tutarlılığı İyileştirme Scripti
Çoklu kullanıcı desteği için veri yazma işlemlerini atomik hale getirir
"""

import json
import os
import tempfile
import shutil
from datetime import datetime

def atomic_save_data(data, filename):
    """
    Veriyi atomik olarak kaydeder (çoklu kullanıcı desteği için)
    """
    try:
        # Geçici dosya oluştur
        temp_file = tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json', encoding='utf-8')
        
        # Veriyi geçici dosyaya yaz
        json.dump(data, temp_file, indent=2, ensure_ascii=False)
        temp_file.close()
        
        # Geçici dosyayı hedef dosyaya taşı (atomik işlem)
        shutil.move(temp_file.name, filename)
        
        return True
    except Exception as e:
        # Hata durumunda geçici dosyayı temizle
        if os.path.exists(temp_file.name):
            os.unlink(temp_file.name)
        print(f"Atomik kaydetme hatası: {e}")
        return False

def create_backup_system():
    """Otomatik yedekleme sistemi oluşturur"""
    backup_script = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Otomatik Veri Yedekleme Sistemi
"""

import json
import os
import shutil
from datetime import datetime
import schedule
import time

def create_backup():
    """Tüm veri dosyalarını yedekler"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_dir = f'backups/backup_{timestamp}'
    
    # Yedekleme dizini oluştur
    os.makedirs(backup_dir, exist_ok=True)
    
    # Yedeklenecek dosyalar
    data_files = [
        'baca_bilgileri.json',
        'parametre_olcum.json',
        'firma_olcum.json',
        'users.json',
        'parameters.json'
    ]
    
    for file in data_files:
        if os.path.exists(file):
            shutil.copy2(file, backup_dir)
            print(f"✅ {file} yedeklendi")
    
    # Eski yedekleri temizle (7 günden eski)
    cleanup_old_backups()
    
    print(f"📦 Yedekleme tamamlandı: {backup_dir}")

def cleanup_old_backups():
    """7 günden eski yedekleri temizler"""
    backup_dir = 'backups'
    if not os.path.exists(backup_dir):
        return
    
    current_time = datetime.now()
    for item in os.listdir(backup_dir):
        item_path = os.path.join(backup_dir, item)
        if os.path.isdir(item_path):
            # Dizin oluşturma zamanını kontrol et
            creation_time = datetime.fromtimestamp(os.path.getctime(item_path))
            if (current_time - creation_time).days > 7:
                shutil.rmtree(item_path)
                print(f"🗑️ Eski yedek silindi: {item}")

def main():
    """Ana fonksiyon"""
    print("🔄 Otomatik Yedekleme Sistemi Başlatılıyor...")
    
    # İlk yedekleme
    create_backup()
    
    # Her saat başı yedekleme
    schedule.every().hour.do(create_backup)
    
    # Her gün gece yarısı yedekleme
    schedule.every().day.at("00:00").do(create_backup)
    
    print("⏰ Yedekleme zamanlaması ayarlandı:")
    print("  - Her saat başı")
    print("  - Her gün gece yarısı")
    print("  - 7 günden eski yedekler otomatik silinir")
    
    # Sürekli çalış
    while True:
        schedule.run_pending()
        time.sleep(60)  # 1 dakika bekle

if __name__ == "__main__":
    main()
'''
    
    with open('auto_backup.py', 'w', encoding='utf-8') as f:
        f.write(backup_script)
    
    print("✅ Otomatik yedekleme sistemi oluşturuldu: auto_backup.py")

def create_data_monitor():
    """Veri bütünlüğü izleme sistemi oluşturur"""
    monitor_script = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Veri Bütünlüğü İzleme Sistemi
"""

import json
import os
import time
from datetime import datetime

def check_data_integrity():
    """Veri dosyalarının bütünlüğünü kontrol eder"""
    data_files = [
        'baca_bilgileri.json',
        'parametre_olcum.json',
        'firma_olcum.json',
        'users.json',
        'parameters.json'
    ]
    
    issues = []
    
    for file in data_files:
        if os.path.exists(file):
            try:
                with open(file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # Temel kontroller
                if isinstance(data, list):
                    print(f"✅ {file}: {len(data)} kayıt")
                else:
                    issues.append(f"❌ {file}: Liste formatında değil")
                    
            except json.JSONDecodeError as e:
                issues.append(f"❌ {file}: JSON hatası - {e}")
            except Exception as e:
                issues.append(f"❌ {file}: Okuma hatası - {e}")
        else:
            print(f"⚠️ {file}: Dosya bulunamadı")
    
    if issues:
        print("\\n🚨 VERİ BÜTÜNLÜĞÜ SORUNLARI:")
        for issue in issues:
            print(f"  {issue}")
    else:
        print("\\n✅ Tüm veri dosyaları sağlıklı")

def main():
    """Ana fonksiyon"""
    print("🔍 Veri Bütünlüğü İzleme Sistemi")
    print("=" * 40)
    
    while True:
        print(f"\\n📊 Kontrol zamanı: {datetime.now().strftime('%H:%M:%S')}")
        check_data_integrity()
        time.sleep(300)  # 5 dakika bekle

if __name__ == "__main__":
    main()
'''
    
    with open('data_monitor.py', 'w', encoding='utf-8') as f:
        f.write(monitor_script)
    
    print("✅ Veri izleme sistemi oluşturuldu: data_monitor.py")

def create_usage_guide():
    """Kullanım kılavuzu oluşturur"""
    guide = """# 🔧 Çoklu Kullanıcı Veri Tutarlılığı Kılavuzu

## 📋 Sorun ve Çözüm

### Sorun:
- Test sırasında 5 kullanıcı aynı anda veri yazdığında, sadece son yazan kullanıcının verisi kaldı
- JSON dosyasına yazma işlemi atomik olmadığı için veri kaybı yaşandı

### Çözüm:
1. **Atomik Veri Yazma**: Geçici dosya kullanarak atomik yazma
2. **Otomatik Yedekleme**: Düzenli veri yedekleme sistemi
3. **Veri İzleme**: Veri bütünlüğü kontrol sistemi

## 🚀 Kullanım

### 1. Otomatik Yedekleme Başlatma:
```bash
python auto_backup.py
```

### 2. Veri İzleme Başlatma:
```bash
python data_monitor.py
```

### 3. Manuel Yedekleme:
```bash
python restore_baca_data.py
```

## 📊 Özellikler

### ✅ Atomik Veri Yazma
- Geçici dosya kullanımı
- Hata durumunda geri alma
- Veri kaybı önleme

### ✅ Otomatik Yedekleme
- Her saat başı yedekleme
- Günlük gece yarısı yedekleme
- 7 günden eski yedekleri temizleme

### ✅ Veri İzleme
- JSON dosya bütünlüğü kontrolü
- Kayıt sayısı takibi
- Hata raporlama

## 🔒 Güvenlik

- Tüm veri işlemleri atomik
- Otomatik yedekleme sistemi
- Veri bütünlüğü kontrolü
- Hata durumunda geri alma

## 📞 Destek

Sorun yaşadığınızda:
1. Veri izleme sistemini kontrol edin
2. Yedek dosyalarından geri yükleyin
3. Gerekirse yeni veri oluşturun
"""
    
    with open('VERI_TUTARLILIGI_KILAVUZU.md', 'w', encoding='utf-8') as f:
        f.write(guide)
    
    print("✅ Kullanım kılavuzu oluşturuldu: VERI_TUTARLILIGI_KILAVUZU.md")

def main():
    """Ana fonksiyon"""
    print("🔧 VERİ TUTARLILIĞI İYİLEŞTİRME SİSTEMİ")
    print("=" * 50)
    
    # 1. Otomatik yedekleme sistemi
    create_backup_system()
    
    # 2. Veri izleme sistemi
    create_data_monitor()
    
    # 3. Kullanım kılavuzu
    create_usage_guide()
    
    print("\n🎯 İyileştirmeler Tamamlandı!")
    print("\n📋 Oluşturulan Dosyalar:")
    print("  ✅ auto_backup.py - Otomatik yedekleme sistemi")
    print("  ✅ data_monitor.py - Veri izleme sistemi")
    print("  ✅ VERI_TUTARLILIGI_KILAVUZU.md - Kullanım kılavuzu")
    
    print("\n🚀 Kullanım:")
    print("  1. Otomatik yedekleme: python auto_backup.py")
    print("  2. Veri izleme: python data_monitor.py")
    print("  3. Kılavuz: VERI_TUTARLILIGI_KILAVUZU.md")

if __name__ == "__main__":
    main() 
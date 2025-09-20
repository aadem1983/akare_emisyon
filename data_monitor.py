#!/usr/bin/env python3
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
        print("\n🚨 VERİ BÜTÜNLÜĞÜ SORUNLARI:")
        for issue in issues:
            print(f"  {issue}")
    else:
        print("\n✅ Tüm veri dosyaları sağlıklı")

def main():
    """Ana fonksiyon"""
    print("🔍 Veri Bütünlüğü İzleme Sistemi")
    print("=" * 40)
    
    while True:
        print(f"\n📊 Kontrol zamanı: {datetime.now().strftime('%H:%M:%S')}")
        check_data_integrity()
        time.sleep(300)  # 5 dakika bekle

if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baca Bilgileri Veri Geri Yükleme Scripti
Test sırasında kaybolan verileri geri yüklemek için
"""

import json
import os
from datetime import datetime

def create_backup():
    """Mevcut veriyi yedekler"""
    if os.path.exists('baca_bilgileri.json'):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_filename = f'baca_bilgileri_backup_{timestamp}.json'
        
        with open('baca_bilgileri.json', 'r', encoding='utf-8') as f:
            current_data = json.load(f)
        
        with open(backup_filename, 'w', encoding='utf-8') as f:
            json.dump(current_data, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Mevcut veri yedeklendi: {backup_filename}")
        return backup_filename
    else:
        print("❌ Yedeklenecek veri bulunamadı")
        return None

def restore_sample_data():
    """Örnek baca bilgileri verisi oluşturur"""
    sample_data = [
        {
            "id": "sample-001",
            "firma_adi": "Örnek Firma A",
            "olcum_kodu": "OLC-2024-001",
            "baca_adi": "Ana Baca",
            "baca_bilgileri": {
                "597fad80-d28f-40ea-bd28-a76c61c5203d": "BACA-001",
                "98399625-5bbc-465e-8e09-de454f231ae4": "Doğalgaz",
                "22867c9a-ca3c-4d80-b017-b73dafdd7fef": "25.5",
                "b1b6fc38-98c0-4048-8b8e-795cf7d44c48": "1.2",
                "8ec9ecc9-ecda-4bf2-9802-02fa2e3fda4c": "30.0",
                "ddca398d-0e55-4662-b661-3731e0975bd2": "Doğalgaz",
                "6b3546e0-184c-49de-82e4-e2835e81923b": "Düz",
                "d9958774-43f7-4bc3-8e12-436614a6193a": "Yuvarlak",
                "6a301d72-f21b-485b-b8fb-116ad5cb223f": "45.0",
                "ab2b67dd-16a5-4bee-9b6e-b60b8cfc2d0c": "5.2",
                "eca60e54-ec39-4412-8884-caa17faed0be": "22.5",
                "64238e0a-6387-4c31-9bf4-d7f800ef17e1": "65.0",
                "b09ad69a-e4d4-4219-b055-2cf923ffd499": "101.3",
                "9c8c8bcf-c98e-4109-8b10-63b08b26460e": "A",
                "af55c55f-f83b-4b90-a655-ee76bf6bb2ac": "B",
                "20881447-f7c8-4a6b-8583-76c7246082ef": "C"
            },
            "personel_adi": "Ahmet Yılmaz",
            "photo_path": None,
            "created_at": "2024-01-15T10:30:00",
            "updated_at": "2024-01-15T10:30:00"
        },
        {
            "id": "sample-002",
            "firma_adi": "Örnek Firma B",
            "olcum_kodu": "OLC-2024-002",
            "baca_adi": "Yan Baca",
            "baca_bilgileri": {
                "597fad80-d28f-40ea-bd28-a76c61c5203d": "BACA-002",
                "98399625-5bbc-465e-8e09-de454f231ae4": "Kömür",
                "22867c9a-ca3c-4d80-b017-b73dafdd7fef": "50.0",
                "b1b6fc38-98c0-4048-8b8e-795cf7d44c48": "1.8",
                "8ec9ecc9-ecda-4bf2-9802-02fa2e3fda4c": "40.0",
                "ddca398d-0e55-4662-b661-3731e0975bd2": "Kömür",
                "6b3546e0-184c-49de-82e4-e2835e81923b": "Eğimli",
                "d9958774-43f7-4bc3-8e12-436614a6193a": "Kare",
                "6a301d72-f21b-485b-b8fb-116ad5cb223f": "60.0",
                "ab2b67dd-16a5-4bee-9b6e-b60b8cfc2d0c": "4.8",
                "eca60e54-ec39-4412-8884-caa17faed0be": "28.0",
                "64238e0a-6387-4c31-9bf4-d7f800ef17e1": "70.0",
                "b09ad69a-e4d4-4219-b055-2cf923ffd499": "101.8",
                "9c8c8bcf-c98e-4109-8b10-63b08b26460e": "A",
                "af55c55f-f83b-4b90-a655-ee76bf6bb2ac": "A",
                "20881447-f7c8-4a6b-8583-76c7246082ef": "B"
            },
            "personel_adi": "Mehmet Demir",
            "photo_path": None,
            "created_at": "2024-01-20T14:15:00",
            "updated_at": "2024-01-20T14:15:00"
        },
        {
            "id": "sample-003",
            "firma_adi": "Örnek Firma C",
            "olcum_kodu": "OLC-2024-003",
            "baca_adi": "Acil Baca",
            "baca_bilgileri": {
                "597fad80-d28f-40ea-bd28-a76c61c5203d": "BACA-003",
                "98399625-5bbc-465e-8e09-de454f231ae4": "Fuel Oil",
                "22867c9a-ca3c-4d80-b017-b73dafdd7fef": "35.0",
                "b1b6fc38-98c0-4048-8b8e-795cf7d44c48": "1.5",
                "8ec9ecc9-ecda-4bf2-9802-02fa2e3fda4c": "35.0",
                "ddca398d-0e55-4662-b661-3731e0975bd2": "Fuel Oil",
                "6b3546e0-184c-49de-82e4-e2835e81923b": "Kubbe",
                "d9958774-43f7-4bc3-8e12-436614a6193a": "Dikdörtgen",
                "6a301d72-f21b-485b-b8fb-116ad5cb223f": "50.0",
                "ab2b67dd-16a5-4bee-9b6e-b60b8cfc2d0c": "6.0",
                "eca60e54-ec39-4412-8884-caa17faed0be": "25.0",
                "64238e0a-6387-4c31-9bf4-d7f800ef17e1": "60.0",
                "b09ad69a-e4d4-4219-b055-2cf923ffd499": "101.5",
                "9c8c8bcf-c98e-4109-8b10-63b08b26460e": "B",
                "af55c55f-f83b-4b90-a655-ee76bf6bb2ac": "C",
                "20881447-f7c8-4a6b-8583-76c7246082ef": "A"
            },
            "personel_adi": "Fatma Kaya",
            "photo_path": None,
            "created_at": "2024-01-25T09:45:00",
            "updated_at": "2024-01-25T09:45:00"
        }
    ]
    
    return sample_data

def main():
    """Ana fonksiyon"""
    print("🔄 BACA BİLGİLERİ VERİ GERİ YÜKLEME")
    print("=" * 50)
    
    # 1. Mevcut veriyi yedekle
    backup_file = create_backup()
    
    # 2. Örnek veri oluştur
    sample_data = restore_sample_data()
    
    # 3. Veriyi kaydet
    with open('baca_bilgileri.json', 'w', encoding='utf-8') as f:
        json.dump(sample_data, f, indent=2, ensure_ascii=False)
    
    print(f"✅ {len(sample_data)} adet örnek baca bilgisi oluşturuldu")
    print("\n📋 Oluşturulan Veriler:")
    for i, data in enumerate(sample_data, 1):
        print(f"  {i}. {data['firma_adi']} - {data['olcum_kodu']} - {data['baca_adi']}")
    
    print(f"\n💾 Veriler 'baca_bilgileri.json' dosyasına kaydedildi")
    if backup_file:
        print(f"📦 Yedek dosya: {backup_file}")

if __name__ == "__main__":
    main() 
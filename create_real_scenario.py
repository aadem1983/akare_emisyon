#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerçek Senaryo Oluşturma Scripti
5 baca ve parametre ölçümleri ile gerçekçi test senaryosu
"""

import json
import os
from datetime import datetime
import uuid

def create_real_firma():
    """Gerçek firma ve baca bilgileri oluşturur"""
    
    # Firma bilgileri
    firma_adi = "AKARE ÇELİK SANAYİ A.Ş."
    olcum_kodu = "AKARE-2024-001"
    
    # 5 baca için bilgiler
    bacalar = [
        {
            "baca_adi": "Ana Üretim Bacası",
            "baca_no": "BACA-001",
            "yakit_turu": "Doğalgaz",
            "isil_guc": "45.5",
            "baca_olcusu": "1.8",
            "yerden_yukseklik": "65.0",
            "catı_yuksekligi": "35.0"
        },
        {
            "baca_adi": "Yan Üretim Bacası",
            "baca_no": "BACA-002", 
            "yakit_turu": "Kömür",
            "isil_guc": "75.0",
            "baca_olcusu": "2.2",
            "yerden_yukseklik": "80.0",
            "catı_yuksekligi": "40.0"
        },
        {
            "baca_adi": "Acil Durum Bacası",
            "baca_no": "BACA-003",
            "yakit_turu": "Fuel Oil",
            "isil_guc": "30.0",
            "baca_olcusu": "1.5",
            "yerden_yukseklik": "55.0",
            "catı_yuksekligi": "30.0"
        },
        {
            "baca_adi": "Yedek Üretim Bacası",
            "baca_no": "BACA-004",
            "yakit_turu": "Doğalgaz",
            "isil_guc": "60.0",
            "baca_olcusu": "2.0",
            "yerden_yukseklik": "70.0",
            "catı_yuksekligi": "38.0"
        },
        {
            "baca_adi": "Test Bacası",
            "baca_no": "BACA-005",
            "yakit_turu": "Kömür",
            "isil_guc": "40.0",
            "baca_olcusu": "1.6",
            "yerden_yukseklik": "60.0",
            "catı_yuksekligi": "32.0"
        }
    ]
    
    # Baca bilgileri oluştur
    baca_bilgileri_list = []
    
    for i, baca in enumerate(bacalar, 1):
        baca_bilgisi = {
            "id": str(uuid.uuid4()),
            "firma_adi": firma_adi,
            "olcum_kodu": olcum_kodu,
            "baca_adi": baca["baca_adi"],
            "kayit_tarihi": datetime.now().strftime('%d.%m.%y'),
            "personel_adi": f"Teknik Ekip {i}",
            "baca_bilgileri": {
                "597fad80-d28f-40ea-bd28-a76c61c5203d": baca["baca_no"],
                "98399625-5bbc-465e-8e09-de454f231ae4": baca["yakit_turu"],
                "22867c9a-ca3c-4d80-b017-b73dafdd7fef": baca["isil_guc"],
                "b1b6fc38-98c0-4048-8b8e-795cf7d44c48": baca["baca_olcusu"],
                "8ec9ecc9-ecda-4bf2-9802-02fa2e3fda4c": baca["catı_yuksekligi"],
                "ddca398d-0e55-4662-b661-3731e0975bd2": baca["yakit_turu"],
                "6b3546e0-184c-49de-82e4-e2835e81923b": "Düz",
                "d9958774-43f7-4bc3-8e12-436614a6193a": "Yuvarlak",
                "6a301d72-f21b-485b-b8fb-116ad5cb223f": baca["yerden_yukseklik"],
                "ab2b67dd-16a5-4bee-9b6e-b60b8cfc2d0c": "5.5",
                "eca60e54-ec39-4412-8884-caa17faed0be": "25.0",
                "64238e0a-6387-4c31-9bf4-d7f800ef17e1": "65.0",
                "b09ad69a-e4d4-4219-b055-2cf923ffd499": "101.3",
                "9c8c8bcf-c98e-4109-8b10-63b08b26460e": "A",
                "af55c55f-f83b-4b90-a655-ee76bf6bb2ac": "B",
                "20881447-f7c8-4a6b-8583-76c7246082ef": "C"
            },
            "photo_path": None,
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat()
        }
        baca_bilgileri_list.append(baca_bilgisi)
    
    return firma_adi, olcum_kodu, bacalar, baca_bilgileri_list

def create_parametre_olcumleri(firma_adi, olcum_kodu, bacalar):
    """Parametre ölçümleri oluşturur"""
    
    parametre_olcumleri = []
    
    # Her baca için 4 parametre (TOZ, YG, TOC, AĞIR.MET)
    parametreler = ["TOZ", "YG", "TOC", "AĞIR.MET"]
    
    for i, baca in enumerate(bacalar, 1):
        for j, parametre in enumerate(parametreler, 1):
            # Rastgele değerler oluştur
            if parametre == "TOZ":
                deger = f"{round(15 + i * 2 + j * 0.5, 1)}"
            elif parametre == "YG":
                deger = f"{round(8 + i * 1.5 + j * 0.3, 1)}"
            elif parametre == "TOC":
                deger = f"{round(3 + i * 0.8 + j * 0.2, 1)}"
            else:  # AĞIR.MET
                deger = f"{round(0.5 + i * 0.1 + j * 0.05, 2)}"
            
            olcum = {
                "id": str(uuid.uuid4()),
                "firma_adi": firma_adi,
                "olcum_kodu": olcum_kodu,
                "baca_adi": baca["baca_adi"],
                "parametre_adi": parametre,
                "parametre_verileri": {
                    "TARİH": datetime.now().strftime('%d.%m.%y'),
                    "METOT": "EPA 5" if parametre == "TOZ" else "EPA 17",
                    "NOZZLE ÇAP": "4.0",
                    "TRAVERS": "12",
                    "B.HIZ": "6.5-6.8-7.0",
                    "B.SIC": "45.0-46.0-47.0",
                    "B.BAS(KPA)": "101.0-101.2-101.3",
                    "B.NEM(G/M3)": "12.5-13.0-13.5",
                    "B.NEM(%)": "50.0-52.0-55.0",
                    "SYC.HAC.": "0.5-0.6-0.7",
                    "SYC.İLK": "0.2-0.25-0.3",
                    "SYC.SON": "0.4-0.45-0.5",
                    "SYC.SIC": "25.0-26.0-27.0",
                    "DEBİ": "100.0-105.0-110.0",
                    "ISDL": "0.2-0.25-0.3",
                    "SONUÇ": deger
                },
                "personel_adi": f"Teknik Ekip {i}",
                "created_at": datetime.now().isoformat(),
                "updated_at": datetime.now().isoformat()
            }
            parametre_olcumleri.append(olcum)
    
    return parametre_olcumleri

def save_data(baca_bilgileri_list, parametre_olcumleri):
    """Verileri dosyalara kaydeder"""
    
    # Baca bilgilerini kaydet
    with open('baca_bilgileri.json', 'w', encoding='utf-8') as f:
        json.dump(baca_bilgileri_list, f, indent=2, ensure_ascii=False)
    
    # Parametre ölçümlerini kaydet
    with open('parametre_olcum.json', 'w', encoding='utf-8') as f:
        json.dump(parametre_olcumleri, f, indent=2, ensure_ascii=False)
    
    print("✅ Veriler başarıyla kaydedildi")

def create_test_users():
    """Test kullanıcıları oluşturur"""
    
    # Mevcut kullanıcıları yükle
    if os.path.exists('users.json'):
        with open('users.json', 'r', encoding='utf-8') as f:
            users = json.load(f)
    else:
        users = {}
    
    # 5 test kullanıcısı ekle
    test_users = {
        "teknik_ekip1": {
            "password": "1111",
            "role": "user",
            "ad_soyad": "Ahmet Yılmaz",
            "gorev": "Teknik Ekip 1 - Ana Üretim Bacası"
        },
        "teknik_ekip2": {
            "password": "1111", 
            "role": "user",
            "ad_soyad": "Mehmet Demir",
            "gorev": "Teknik Ekip 2 - Yan Üretim Bacası"
        },
        "teknik_ekip3": {
            "password": "1111",
            "role": "user", 
            "ad_soyad": "Fatma Kaya",
            "gorev": "Teknik Ekip 3 - Acil Durum Bacası"
        },
        "teknik_ekip4": {
            "password": "1111",
            "role": "user",
            "ad_soyad": "Ali Özkan",
            "gorev": "Teknik Ekip 4 - Yedek Üretim Bacası"
        },
        "teknik_ekip5": {
            "password": "1111",
            "role": "user",
            "ad_soyad": "Ayşe Çelik",
            "gorev": "Teknik Ekip 5 - Test Bacası"
        }
    }
    
    # Yeni kullanıcıları ekle
    users.update(test_users)
    
    # Kullanıcıları kaydet
    with open('users.json', 'w', encoding='utf-8') as f:
        json.dump(users, f, indent=2, ensure_ascii=False)
    
    print("✅ Test kullanıcıları oluşturuldu")

def main():
    """Ana fonksiyon"""
    print("🏭 GERÇEK SENARYO OLUŞTURMA")
    print("=" * 50)
    
    # 1. Firma ve baca bilgileri oluştur
    firma_adi, olcum_kodu, bacalar, baca_bilgileri_list = create_real_firma()
    
    print(f"🏢 Firma: {firma_adi}")
    print(f"📋 Ölçüm Kodu: {olcum_kodu}")
    print(f"🏗️ Baca Sayısı: {len(bacalar)}")
    
    print("\n📊 Baca Listesi:")
    for i, baca in enumerate(bacalar, 1):
        print(f"  {i}. {baca['baca_adi']} ({baca['baca_no']})")
    
    # 2. Parametre ölçümleri oluştur
    parametre_olcumleri = create_parametre_olcumleri(firma_adi, olcum_kodu, bacalar)
    
    print(f"\n🔬 Parametre Ölçümleri:")
    print(f"  Toplam: {len(parametre_olcumleri)} ölçüm")
    print(f"  Her baca için: TOZ, YG, TOC, AĞIR.MET")
    
    # 3. Test kullanıcıları oluştur
    create_test_users()
    
    # 4. Verileri kaydet
    save_data(baca_bilgileri_list, parametre_olcumleri)
    
    print("\n🎯 SENARYO HAZIR!")
    print("\n👥 Test Kullanıcıları:")
    print("  teknik_ekip1 - Ana Üretim Bacası (BACA-001)")
    print("  teknik_ekip2 - Yan Üretim Bacası (BACA-002)")
    print("  teknik_ekip3 - Acil Durum Bacası (BACA-003)")
    print("  teknik_ekip4 - Yedek Üretim Bacası (BACA-004)")
    print("  teknik_ekip5 - Test Bacası (BACA-005)")
    
    print("\n📋 Her kullanıcı kendi baca numarasına ait parametreleri görecek:")
    print("  - TOZ (mg/m³)")
    print("  - YG (mg/m³)")
    print("  - TOC (mg/m³)")
    print("  - AĞIR.MET (mg/m³)")
    
    print(f"\n🚀 Test için: python real_scenario_test.py")

if __name__ == "__main__":
    main() 
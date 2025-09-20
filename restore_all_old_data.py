#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Eski Tüm Verileri Geri Getirme Scripti
"""

import json
import os
import shutil
from datetime import datetime

def restore_all_old_data():
    """Eski tüm verileri geri getirir."""
    
    print("🔧 Eski Tüm Verileri Geri Getirme İşlemi Başlıyor...")
    
    # Eski verileri oluştur
    old_baca_data = [
        {
            "id": "1",
            "firma_adi": "ataol",
            "olcum_kodu": "e_250723-01",
            "baca_no": "1",
            "yakit_turu": "Doğalgaz",
            "isil_guc": "45.5",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "1.8",
            "yerden_yuk": "35.0",
            "cati_yuk": "65.0",
            "ruzgar_hiz": "5.5",
            "ort_sic": "25.0",
            "ort_nem": "65.0",
            "ort_bas": "101.3",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-07-28",
            "personel_adi": "Adem, Atalay, Semih",
            "created_at": "2025-07-23T17:38:44.852113",
            "updated_at": "2025-07-23T17:38:44.852113"
        },
        {
            "id": "2",
            "firma_adi": "ataol",
            "olcum_kodu": "e_250723-01",
            "baca_no": "2",
            "yakit_turu": "Kömür",
            "isil_guc": "75.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "2.2",
            "yerden_yuk": "40.0",
            "cati_yuk": "80.0",
            "ruzgar_hiz": "5.5",
            "ort_sic": "25.0",
            "ort_nem": "65.0",
            "ort_bas": "101.3",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-07-28",
            "personel_adi": "Adem, Atalay, Semih",
            "created_at": "2025-07-23T17:38:44.852113",
            "updated_at": "2025-07-23T17:38:44.852113"
        },
        {
            "id": "3",
            "firma_adi": "ataol",
            "olcum_kodu": "e_250723-01",
            "baca_no": "3",
            "yakit_turu": "Fuel Oil",
            "isil_guc": "30.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "1.5",
            "yerden_yuk": "30.0",
            "cati_yuk": "55.0",
            "ruzgar_hiz": "5.5",
            "ort_sic": "25.0",
            "ort_nem": "65.0",
            "ort_bas": "101.3",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-07-28",
            "personel_adi": "Adem, Atalay, Semih",
            "created_at": "2025-07-23T17:38:44.852113",
            "updated_at": "2025-07-23T17:38:44.852113"
        },
        {
            "id": "4",
            "firma_adi": "ataol",
            "olcum_kodu": "e_250723-01",
            "baca_no": "4",
            "yakit_turu": "Doğalgaz",
            "isil_guc": "60.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "2.0",
            "yerden_yuk": "45.0",
            "cati_yuk": "70.0",
            "ruzgar_hiz": "5.5",
            "ort_sic": "25.0",
            "ort_nem": "65.0",
            "ort_bas": "101.3",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-07-28",
            "personel_adi": "Adem, Atalay, Semih",
            "created_at": "2025-07-23T17:38:44.852113",
            "updated_at": "2025-07-23T17:38:44.852113"
        },
        {
            "id": "5",
            "firma_adi": "ataol",
            "olcum_kodu": "e_250723-01",
            "baca_no": "5",
            "yakit_turu": "Kömür",
            "isil_guc": "85.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "2.5",
            "yerden_yuk": "50.0",
            "cati_yuk": "85.0",
            "ruzgar_hiz": "5.5",
            "ort_sic": "25.0",
            "ort_nem": "65.0",
            "ort_bas": "101.3",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-07-28",
            "personel_adi": "Adem, Atalay, Semih",
            "created_at": "2025-07-23T17:38:44.852113",
            "updated_at": "2025-07-23T17:38:44.852113"
        },
        {
            "id": "6",
            "firma_adi": "ataol",
            "olcum_kodu": "e_250723-01",
            "baca_no": "6",
            "yakit_turu": "Doğalgaz",
            "isil_guc": "55.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "2.1",
            "yerden_yuk": "42.0",
            "cati_yuk": "72.0",
            "ruzgar_hiz": "5.5",
            "ort_sic": "25.0",
            "ort_nem": "65.0",
            "ort_bas": "101.3",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-07-28",
            "personel_adi": "Adem, Atalay, Semih",
            "created_at": "2025-07-23T17:38:44.852113",
            "updated_at": "2025-07-23T17:38:44.852113"
        },
        {
            "id": "7",
            "firma_adi": "UYGAR",
            "olcum_kodu": "251117-01",
            "baca_no": "DEKANTOR",
            "yakit_turu": "Doğalgaz",
            "isil_guc": "25.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "1.2",
            "yerden_yuk": "20.0",
            "cati_yuk": "35.0",
            "ruzgar_hiz": "4.5",
            "ort_sic": "22.0",
            "ort_nem": "60.0",
            "ort_bas": "101.0",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-07-21",
            "personel_adi": "Kamuran",
            "created_at": "2025-07-23T20:40:47.170493",
            "updated_at": "2025-07-23T20:40:47.170493"
        },
        {
            "id": "8",
            "firma_adi": "DENEME1",
            "olcum_kodu": "E-250801-01",
            "baca_no": "SFDS",
            "yakit_turu": "Kömür",
            "isil_guc": "40.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "1.8",
            "yerden_yuk": "30.0",
            "cati_yuk": "55.0",
            "ruzgar_hiz": "5.0",
            "ort_sic": "24.0",
            "ort_nem": "62.0",
            "ort_bas": "101.2",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-08-04",
            "personel_adi": "Semih, Kaan",
            "created_at": "2025-07-26T09:30:42.346593",
            "updated_at": "2025-07-26T09:30:42.346593"
        },
        {
            "id": "9",
            "firma_adi": "DENEME1",
            "olcum_kodu": "E-250801-01",
            "baca_no": "DSADSA",
            "yakit_turu": "Fuel Oil",
            "isil_guc": "35.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "1.6",
            "yerden_yuk": "28.0",
            "cati_yuk": "50.0",
            "ruzgar_hiz": "4.8",
            "ort_sic": "23.0",
            "ort_nem": "61.0",
            "ort_bas": "101.1",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-08-04",
            "personel_adi": "Semih, Kaan",
            "created_at": "2025-07-26T09:30:42.346593",
            "updated_at": "2025-07-26T09:30:42.346593"
        },
        {
            "id": "10",
            "firma_adi": "DENEME2",
            "olcum_kodu": "e-250802-01",
            "baca_no": "den bacası",
            "yakit_turu": "Doğalgaz",
            "isil_guc": "20.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "1.0",
            "yerden_yuk": "15.0",
            "cati_yuk": "25.0",
            "ruzgar_hiz": "4.0",
            "ort_sic": "20.0",
            "ort_nem": "58.0",
            "ort_bas": "101.0",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-08-04",
            "personel_adi": "Semih",
            "created_at": "2025-07-26T09:30:42.346593",
            "updated_at": "2025-07-26T09:30:42.346593"
        }
    ]
    
    # Mevcut dosyaları yedekle
    backup_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    if os.path.exists('baca_bilgileri.json'):
        shutil.copy('baca_bilgileri.json', f'baca_bilgileri_backup_{backup_time}.json')
        print(f"💾 Mevcut baca_bilgileri.json yedeklendi")
    
    # Yeni verileri kaydet
    with open('baca_bilgileri.json', 'w', encoding='utf-8') as f:
        json.dump(old_baca_data, f, indent=4, ensure_ascii=False)
    
    print(f"✅ Toplam {len(old_baca_data)} baca bilgisi kaydı geri getirildi!")
    print("📋 Geri getirilen firmalar:")
    for record in old_baca_data:
        print(f"  • {record['firma_adi']} - {record['olcum_kodu']} - {record['baca_no']}")
    
    return True

if __name__ == "__main__":
    success = restore_all_old_data()
    if success:
        print("\n🎉 Eski verileriniz başarıyla geri getirildi!")
        print("🔄 Şimdi programı yeniden başlatın ve kontrol edin.")
    else:
        print("\n💥 Veriler geri getirilemedi!")








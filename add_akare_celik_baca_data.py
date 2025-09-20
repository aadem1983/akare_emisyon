#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AKARE ÇELİK Baca Bilgilerini Ekleme Scripti
"""

import json
import os
from datetime import datetime

def add_akare_celik_baca_data():
    """AKARE ÇELİK firmasının baca bilgilerini ekler."""
    
    print("🔧 AKARE ÇELİK Baca Bilgileri Ekleme İşlemi Başlıyor...")
    
    # Mevcut baca bilgilerini yükle
    with open('baca_bilgileri.json', 'r', encoding='utf-8') as f:
        current_data = json.load(f)
    
    # AKARE ÇELİK baca bilgileri
    akare_celik_bacalar = [
        {
            "id": "akare_1",
            "firma_adi": "AKARE ÇELİK SANAYİ A.Ş.",
            "olcum_kodu": "AKARE-2024-001",
            "baca_no": "Ana Üretim Bacası",
            "yakit_turu": "Doğalgaz",
            "isil_guc": "45.5",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "1.8",
            "yerden_yuk": "35.0",
            "cati_yuk": "65.0",
            "ruzgar_hiz": "6.5",
            "ort_sic": "45.0",
            "ort_nem": "12.5",
            "ort_bas": "101.0",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-08-07",
            "personel_adi": "Teknik Ekip 1",
            "created_at": "2025-08-07T01:50:32.817173",
            "updated_at": "2025-08-07T01:50:32.817189"
        },
        {
            "id": "akare_2",
            "firma_adi": "AKARE ÇELİK SANAYİ A.Ş.",
            "olcum_kodu": "AKARE-2024-001",
            "baca_no": "Yan Üretim Bacası",
            "yakit_turu": "Kömür",
            "isil_guc": "75.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "2.2",
            "yerden_yuk": "40.0",
            "cati_yuk": "80.0",
            "ruzgar_hiz": "6.3",
            "ort_sic": "47.6",
            "ort_nem": "10.3",
            "ort_bas": "100.5",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-08-07",
            "personel_adi": "Teknik Ekip 2",
            "created_at": "2025-08-07T01:50:32.817259",
            "updated_at": "2025-08-07T01:50:32.817265"
        },
        {
            "id": "akare_3",
            "firma_adi": "AKARE ÇELİK SANAYİ A.Ş.",
            "olcum_kodu": "AKARE-2024-001",
            "baca_no": "Acil Durum Bacası",
            "yakit_turu": "Fuel Oil",
            "isil_guc": "30.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "1.5",
            "yerden_yuk": "30.0",
            "cati_yuk": "55.0",
            "ruzgar_hiz": "6.5",
            "ort_sic": "45.0",
            "ort_nem": "12.5",
            "ort_bas": "101.0",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-08-07",
            "personel_adi": "Teknik Ekip 3",
            "created_at": "2025-08-07T01:50:32.817388",
            "updated_at": "2025-08-07T01:50:32.817394"
        },
        {
            "id": "akare_4",
            "firma_adi": "AKARE ÇELİK SANAYİ A.Ş.",
            "olcum_kodu": "AKARE-2024-001",
            "baca_no": "Yedek Üretim Bacası",
            "yakit_turu": "Doğalgaz",
            "isil_guc": "60.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "2.0",
            "yerden_yuk": "45.0",
            "cati_yuk": "70.0",
            "ruzgar_hiz": "6.5",
            "ort_sic": "45.0",
            "ort_nem": "12.5",
            "ort_bas": "101.0",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-08-07",
            "personel_adi": "Teknik Ekip 4",
            "created_at": "2025-08-07T01:50:32.817388",
            "updated_at": "2025-08-07T01:50:32.817394"
        },
        {
            "id": "akare_5",
            "firma_adi": "AKARE ÇELİK SANAYİ A.Ş.",
            "olcum_kodu": "AKARE-2024-001",
            "baca_no": "Test Bacası",
            "yakit_turu": "Doğalgaz",
            "isil_guc": "25.0",
            "cati_sekli": "Düz",
            "kaynak_turu": "PROSES",
            "baca_sekli": "Yuvarlak",
            "baca_olcusu": "1.2",
            "yerden_yuk": "20.0",
            "cati_yuk": "35.0",
            "ruzgar_hiz": "6.5",
            "ort_sic": "45.0",
            "ort_nem": "12.5",
            "ort_bas": "101.0",
            "a_baca": "A",
            "b_baca": "B",
            "c_delik": "C",
            "foto": None,
            "kayit_tarihi": "2025-08-07",
            "personel_adi": "Teknik Ekip 5",
            "created_at": "2025-08-07T01:50:32.817388",
            "updated_at": "2025-08-07T01:50:32.817394"
        }
    ]
    
    # AKARE ÇELİK bacalarını mevcut listeye ekle
    current_data.extend(akare_celik_bacalar)
    
    # Güncellenmiş verileri kaydet
    with open('baca_bilgileri.json', 'w', encoding='utf-8') as f:
        json.dump(current_data, f, indent=4, ensure_ascii=False)
    
    print(f"✅ AKARE ÇELİK için {len(akare_celik_bacalar)} baca bilgisi eklendi!")
    print("📋 Eklenen AKARE ÇELİK bacaları:")
    for baca in akare_celik_bacalar:
        print(f"  • {baca['baca_no']}")
    
    print(f"📊 Toplam baca sayısı: {len(current_data)}")
    
    return True

if __name__ == "__main__":
    success = add_akare_celik_baca_data()
    if success:
        print("\n🎉 AKARE ÇELİK baca bilgileri başarıyla eklendi!")
        print("🔄 Şimdi programı yeniden başlatın ve kontrol edin.")
    else:
        print("\n💥 AKARE ÇELİK baca bilgileri eklenemedi!")








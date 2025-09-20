#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Senaryo Verilerini Firma Ölçüm Dosyasına Ekleme Scripti
AKARE ÇELİK SANAYİ A.Ş. verilerini firma_olcum.json'a ekler
"""

import json
import os
from datetime import datetime
import uuid

def add_scenario_to_firma_olcum():
    """Senaryo verilerini firma_olcum.json dosyasına ekler"""
    
    # Mevcut firma ölçüm verilerini yükle
    if os.path.exists('firma_olcum.json'):
        with open('firma_olcum.json', 'r', encoding='utf-8') as f:
            firma_olcumler = json.load(f)
    else:
        firma_olcumler = []
    
    # Senaryo verilerini oluştur
    scenario_data = {
        "id": str(uuid.uuid4()),
        "firma_adi": "AKARE ÇELİK SANAYİ A.Ş.",
        "olcum_kodu": "AKARE-2024-001",
        "baslangic_tarihi": "2025-08-07",
        "bitis_tarihi": "2025-08-10",
        "il": "KOCAELİ",
        "ilce": "İZMİT",
        "yetkili": "Teknik Müdür",
        "telefon": "0555-123-4567",
        "durum": "Aktif",
        "personel": [
            "Ahmet Yılmaz",
            "Mehmet Demir", 
            "Fatma Kaya",
            "Ali Özkan",
            "Ayşe Çelik"
        ],
        "baca_sayisi": "5",
        "baca_parametreleri": {
            "Ana Üretim Bacası": ["TOZ", "YG", "TOC", "AĞIR.MET"],
            "Yan Üretim Bacası": ["TOZ", "YG", "TOC", "AĞIR.MET"],
            "Acil Durum Bacası": ["TOZ", "YG", "TOC", "AĞIR.MET"],
            "Yedek Üretim Bacası": ["TOZ", "YG", "TOC", "AĞIR.MET"],
            "Test Bacası": ["TOZ", "YG", "TOC", "AĞIR.MET"]
        },
        "notlar": "Gerçek senaryo test verileri - 5 baca ve 4 parametre ölçümü",
        "olusturma_tarihi": datetime.now().isoformat()
    }
    
    # Aynı firma ve ölçüm kodu varsa güncelle, yoksa ekle
    existing_index = None
    for i, olcum in enumerate(firma_olcumler):
        if (olcum.get('firma_adi') == scenario_data['firma_adi'] and 
            olcum.get('olcum_kodu') == scenario_data['olcum_kodu']):
            existing_index = i
            break
    
    if existing_index is not None:
        # Mevcut kaydı güncelle
        firma_olcumler[existing_index] = scenario_data
        print("✅ Mevcut kayıt güncellendi")
    else:
        # Yeni kayıt ekle
        firma_olcumler.append(scenario_data)
        print("✅ Yeni kayıt eklendi")
    
    # Dosyayı kaydet
    with open('firma_olcum.json', 'w', encoding='utf-8') as f:
        json.dump(firma_olcumler, f, indent=2, ensure_ascii=False)
    
    print(f"✅ Firma ölçüm verileri kaydedildi")
    print(f"📊 Toplam kayıt sayısı: {len(firma_olcumler)}")
    
    return scenario_data

def verify_data():
    """Verilerin doğru eklendiğini kontrol eder"""
    
    # Baca bilgilerini kontrol et
    if os.path.exists('baca_bilgileri.json'):
        with open('baca_bilgileri.json', 'r', encoding='utf-8') as f:
            baca_bilgileri = json.load(f)
        
        akare_bacalar = [b for b in baca_bilgileri if b.get('firma_adi') == "AKARE ÇELİK SANAYİ A.Ş."]
        print(f"🏗️ Baca bilgileri: {len(akare_bacalar)} kayıt")
    
    # Parametre ölçümlerini kontrol et
    if os.path.exists('parametre_olcum.json'):
        with open('parametre_olcum.json', 'r', encoding='utf-8') as f:
            parametre_olcumleri = json.load(f)
        
        akare_parametreler = [p for p in parametre_olcumleri if p.get('firma_adi') == "AKARE ÇELİK SANAYİ A.Ş."]
        print(f"🔬 Parametre ölçümleri: {len(akare_parametreler)} kayıt")
    
    # Firma ölçümlerini kontrol et
    if os.path.exists('firma_olcum.json'):
        with open('firma_olcum.json', 'r', encoding='utf-8') as f:
            firma_olcumler = json.load(f)
        
        akare_firma = [f for f in firma_olcumler if f.get('firma_adi') == "AKARE ÇELİK SANAYİ A.Ş."]
        print(f"🏢 Firma ölçümleri: {len(akare_firma)} kayıt")
        
        if akare_firma:
            firma = akare_firma[0]
            print(f"   Firma: {firma.get('firma_adi')}")
            print(f"   Ölçüm Kodu: {firma.get('olcum_kodu')}")
            print(f"   Baca Sayısı: {firma.get('baca_sayisi')}")
            print(f"   Personel: {', '.join(firma.get('personel', []))}")

def main():
    """Ana fonksiyon"""
    print("🏭 SENARYO VERİLERİNİ FIRMA ÖLÇÜM DOSYASINA EKLEME")
    print("=" * 60)
    
    # Senaryo verilerini ekle
    scenario_data = add_scenario_to_firma_olcum()
    
    print(f"\n📋 Eklenen Veriler:")
    print(f"   Firma: {scenario_data['firma_adi']}")
    print(f"   Ölçüm Kodu: {scenario_data['olcum_kodu']}")
    print(f"   Baca Sayısı: {scenario_data['baca_sayisi']}")
    print(f"   Parametreler: TOZ, YG, TOC, AĞIR.MET")
    
    print(f"\n🔍 Veri Kontrolü:")
    verify_data()
    
    print(f"\n✅ İşlem tamamlandı!")
    print(f"🎯 Artık firma genel raporunda AKARE ÇELİK SANAYİ A.Ş. verilerini görebilirsiniz.")

if __name__ == "__main__":
    main() 
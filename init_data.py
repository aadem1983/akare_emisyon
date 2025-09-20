#!/usr/bin/env python3
"""
Fly.io için veri dosyalarını başlatma scripti
Bu script, gerekli JSON dosyalarının mevcut olup olmadığını kontrol eder
ve yoksa boş dosyalar oluşturur.
"""

import os
import json

def init_data_files():
    """Gerekli veri dosyalarını başlat"""
    
    # Veri dosyaları listesi
    data_files = [
        'firma_kayit.json',
        'firma_olcum.json', 
        'baca_bilgileri.json',
        'parametre_olcum.json',
        'parametre_sahabil.json',
        'saha_olc.json',
        'saha.json',
        'parametre_fields.json',
        'teklif.json',
        'forms.json',
        'rapor_sablonu.json',
        'measurements.json',
        'parametre_olcum.json',
        'users.json',
        'used_teklif_numbers.json'
    ]
    
    print("Fly.io veri dosyaları başlatılıyor...")
    
    for filename in data_files:
        if not os.path.exists(filename):
            print(f"Oluşturuluyor: {filename}")
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False, indent=2)
        else:
            print(f"Mevcut: {filename}")
    
    print("Veri dosyaları hazır!")

if __name__ == "__main__":
    init_data_files()

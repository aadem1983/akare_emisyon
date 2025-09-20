#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baca Bilgileri Veri Formatı Düzeltme Scripti
Eski formattan yeni formata dönüştürür
"""

import json
import os
from datetime import datetime

def fix_baca_data_format():
    """Baca bilgileri verilerini eski formattan yeni formata dönüştürür."""
    
    # Dosya yolları
    baca_file = 'baca_bilgileri.json'
    backup_file = f'baca_bilgileri_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
    
    print("🔧 Baca Bilgileri Veri Formatı Düzeltme İşlemi Başlıyor...")
    
    try:
        # Mevcut veriyi yükle
        if not os.path.exists(baca_file):
            print("❌ baca_bilgileri.json dosyası bulunamadı!")
            return False
            
        with open(baca_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        print(f"📊 Toplam {len(data)} kayıt bulundu")
        
        # Yedek oluştur
        with open(backup_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        print(f"💾 Yedek oluşturuldu: {backup_file}")
        
        # Yeni format verileri
        new_data = []
        
        for item in data:
            # Eski format kontrolü
            if 'baca_bilgileri' in item and isinstance(item['baca_bilgileri'], dict):
                print(f"🔄 Eski format kayıt dönüştürülüyor: {item.get('baca_adi', 'Bilinmeyen')}")
                
                # Yeni format oluştur
                new_item = {
                    'id': item.get('id', ''),
                    'firma_adi': item.get('firma_adi', ''),
                    'olcum_kodu': item.get('olcum_kodu', ''),
                    'baca_no': item.get('baca_adi', ''),  # baca_adi -> baca_no
                    'yakit_turu': item.get('baca_bilgileri', {}).get('98399625-5bbc-465e-8e09-de454f231ae4', ''),
                    'isil_guc': item.get('baca_bilgileri', {}).get('22867c9a-ca3c-4d80-b017-b73dafdd7fef', ''),
                    'cati_sekli': item.get('baca_bilgileri', {}).get('6b3546e0-184c-49de-82e4-e2835e81923b', ''),
                    'kaynak_turu': item.get('baca_bilgileri', {}).get('ddca398d-0e55-4662-b661-3731e0975bd2', ''),
                    'baca_sekli': item.get('baca_bilgileri', {}).get('d9958774-43f7-4bc3-8e12-436614a6193a', ''),
                    'baca_olcusu': item.get('baca_bilgileri', {}).get('b1b6fc38-98c0-4048-8b8e-795cf7d44c48', ''),
                    'yerden_yuk': item.get('baca_bilgileri', {}).get('8ec9ecc9-ecda-4bf2-9802-02fa2e3fda4c', ''),
                    'cati_yuk': item.get('baca_bilgileri', {}).get('6a301d72-f21b-485b-b8fb-116ad5cb223f', ''),
                    'ruzgar_hiz': item.get('baca_bilgileri', {}).get('ab2b67dd-16a5-4bee-9b6e-b60b8cfc2d0c', ''),
                    'ort_sic': item.get('baca_bilgileri', {}).get('eca60e54-ec39-4412-8884-caa17faed0be', ''),
                    'ort_nem': item.get('baca_bilgileri', {}).get('64238e0a-6387-4c31-9bf4-d7f800ef17e1', ''),
                    'ort_bas': item.get('baca_bilgileri', {}).get('b09ad69a-e4d4-4219-b055-2cf923ffd499', ''),
                    'a_baca': item.get('baca_bilgileri', {}).get('9c8c8bcf-c98e-4109-8b10-63b08b26460e', ''),
                    'b_baca': item.get('baca_bilgileri', {}).get('af55c55f-f83b-4b90-a655-ee76bf6bb2ac', ''),
                    'c_delik': item.get('baca_bilgileri', {}).get('20881447-f7c8-4a6b-8583-76c7246082ef', ''),
                    'foto': item.get('photo_path', ''),
                    'kayit_tarihi': item.get('kayit_tarihi', ''),
                    'personel_adi': item.get('personel_adi', ''),
                    'created_at': item.get('created_at', ''),
                    'updated_at': item.get('updated_at', '')
                }
                
                new_data.append(new_item)
                print(f"✅ Dönüştürüldü: {new_item['baca_no']}")
                
            else:
                # Zaten yeni format, olduğu gibi ekle
                print(f"ℹ️ Zaten yeni format: {item.get('baca_no', 'Bilinmeyen')}")
                new_data.append(item)
        
        # Yeni veriyi kaydet
        with open(baca_file, 'w', encoding='utf-8') as f:
            json.dump(new_data, f, indent=4, ensure_ascii=False)
        
        print(f"✅ Toplam {len(new_data)} kayıt yeni formata dönüştürüldü ve kaydedildi!")
        print(f"📁 Dosya: {baca_file}")
        
        return True
        
    except Exception as e:
        print(f"❌ Hata oluştu: {e}")
        return False

if __name__ == "__main__":
    success = fix_baca_data_format()
    if success:
        print("\n🎉 Veri formatı düzeltme işlemi başarıyla tamamlandı!")
        print("🔄 Şimdi programı yeniden başlatın ve Baca Bilgileri sayfasını kontrol edin.")
    else:
        print("\n💥 Veri formatı düzeltme işlemi başarısız!") 
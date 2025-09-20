#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Eski Baca Bilgileri Verilerini Geri Getirme Scripti
Firma ölçüm verilerinden baca bilgileri oluşturur
"""

import json
import os
from datetime import datetime
from uuid import uuid4

def restore_old_baca_data():
    """Eski baca bilgileri verilerini geri getirir."""
    
    print("🔧 Eski Baca Bilgileri Verilerini Geri Getirme İşlemi Başlıyor...")
    
    try:
        # Firma ölçüm verilerini yükle
        firma_olcum_file = 'firma_olcum.json'
        if not os.path.exists(firma_olcum_file):
            print("❌ firma_olcum.json dosyası bulunamadı!")
            return False
            
        with open(firma_olcum_file, 'r', encoding='utf-8') as f:
            firma_olcum_data = json.load(f)
        
        print(f"📊 Toplam {len(firma_olcum_data)} firma ölçüm kaydı bulundu")
        
        # Yeni baca bilgileri listesi
        new_baca_data = []
        
        # Her firma ölçümü için baca bilgileri oluştur
        for firma_olcum in firma_olcum_data:
            firma_adi = firma_olcum.get('firma_adi', '')
            olcum_kodu = firma_olcum.get('olcum_kodu', '')
            baca_parametreleri = firma_olcum.get('baca_parametreleri', {})
            
            print(f"🔄 İşleniyor: {firma_adi} - {olcum_kodu}")
            
            # Her baca için kayıt oluştur
            for baca_adi, parametreler in baca_parametreleri.items():
                if baca_adi:  # Boş baca adı değilse
                    print(f"  📝 Baca: {baca_adi}")
                    
                    # Baca bilgisi oluştur
                    baca_record = {
                        'id': str(uuid4()),
                        'firma_adi': firma_adi,
                        'olcum_kodu': olcum_kodu,
                        'baca_no': baca_adi,
                        'yakit_turu': 'Belirtilmemiş',
                        'isil_guc': '',
                        'cati_sekli': 'Düz',
                        'kaynak_turu': 'PROSES',
                        'baca_sekli': 'Yuvarlak',
                        'baca_olcusu': '',
                        'yerden_yuk': '',
                        'cati_yuk': '',
                        'ruzgar_hiz': '',
                        'ort_sic': '',
                        'ort_nem': '',
                        'ort_bas': '',
                        'a_baca': '',
                        'b_baca': '',
                        'c_delik': '',
                        'foto': None,
                        'kayit_tarihi': firma_olcum.get('baslangic_tarihi', ''),
                        'personel_adi': ', '.join(firma_olcum.get('personel', [])),
                        'created_at': firma_olcum.get('olusturma_tarihi', ''),
                        'updated_at': firma_olcum.get('olusturma_tarihi', '')
                    }
                    
                    new_baca_data.append(baca_record)
        
        # Mevcut baca bilgileri dosyasını yedekle
        current_baca_file = 'baca_bilgileri.json'
        if os.path.exists(current_baca_file):
            backup_file = f'baca_bilgileri_current_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
            with open(backup_file, 'w', encoding='utf-8') as f:
                with open(current_baca_file, 'r', encoding='utf-8') as current_f:
                    f.write(current_f.read())
            print(f"💾 Mevcut veriler yedeklendi: {backup_file}")
        
        # Yeni verileri kaydet
        with open(current_baca_file, 'w', encoding='utf-8') as f:
            json.dump(new_baca_data, f, indent=4, ensure_ascii=False)
        
        print(f"✅ Toplam {len(new_baca_data)} baca bilgisi kaydı oluşturuldu ve kaydedildi!")
        print(f"📁 Dosya: {current_baca_file}")
        
        # Özet bilgi
        print("\n📋 Oluşturulan Kayıtlar:")
        for record in new_baca_data:
            print(f"  • {record['firma_adi']} - {record['olcum_kodu']} - {record['baca_no']}")
        
        return True
        
    except Exception as e:
        print(f"❌ Hata oluştu: {e}")
        return False

if __name__ == "__main__":
    success = restore_old_baca_data()
    if success:
        print("\n🎉 Eski baca bilgileri verileri başarıyla geri getirildi!")
        print("🔄 Şimdi programı yeniden başlatın ve Baca Bilgileri sayfasını kontrol edin.")
    else:
        print("\n💥 Eski veriler geri getirilemedi!") 
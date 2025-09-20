#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Firebase Veri Migrasyon Scripti - Akare Emisyon Saha Uygulaması
JSON dosyalarından Firestore'a veri aktarımı
"""

import json
import firebase_admin
from firebase_admin import credentials, firestore
import os
from datetime import datetime

def load_json_file(filename):
    """JSON dosyasını yükle"""
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            data = json.load(f)
            print(f"✅ {filename} yüklendi - {len(data) if isinstance(data, (list, dict)) else 'Tek kayıt'} kayıt")
            return data
    except FileNotFoundError:
        print(f"❌ {filename} bulunamadı")
        return None
    except json.JSONDecodeError as e:
        print(f"❌ {filename} JSON hatası: {e}")
        return None

def migrate_users(db, users_data):
    """Kullanıcıları migrate et"""
    if not users_data:
        print("⚠️ Kullanıcı verisi bulunamadı")
        return
    
    users_ref = db.collection('users')
    migrated_count = 0
    
    try:
        for username, user_data in users_data.items():
            # Kullanıcı ID'si olarak username kullan
            users_ref.document(username).set(user_data)
            migrated_count += 1
            print(f"  📝 Kullanıcı eklendi: {username}")
        
        print(f"✅ {migrated_count} kullanıcı migrate edildi")
    except Exception as e:
        print(f"❌ Kullanıcı migrasyon hatası: {e}")

def migrate_parameters(db, parameters_data):
    """Parametreleri migrate et"""
    if not parameters_data:
        print("⚠️ Parametre verisi bulunamadı")
        return
    
    parameters_ref = db.collection('parameters')
    migrated_count = 0
    
    try:
        # Parametreler liste formatında
        if isinstance(parameters_data, list):
            for param in parameters_data:
                # JSON formatındaki alanları Firestore formatına çevir
                firestore_data = {
                    'parametre_adi': param.get('Parametre Adı', ''),
                    'metot': param.get('Metot', ''),
                    'izo_oran': param.get('İzo Oran', ''),
                    'nozzle': param.get('Nozzle', ''),
                    'imp1': param.get('1. İmp', ''),
                    'imp2': param.get('2. İmp', ''),
                    'imp3': param.get('3. İmp', ''),
                    'imp4': param.get('4. İmp', ''),
                    'l_dak': param.get('L/DAK', ''),
                    't_hac': param.get('T.HAC', ''),
                    'loq': param.get('LOQ', ''),
                    'id': param.get('id', ''),
                    'created_at': firestore.SERVER_TIMESTAMP
                }
                
                # ID varsa kullan, yoksa otomatik oluştur
                if param.get('id'):
                    parameters_ref.document(param['id']).set(firestore_data)
                else:
                    parameters_ref.add(firestore_data)
                
                migrated_count += 1
                print(f"  📝 Parametre eklendi: {param.get('Parametre Adı', 'Bilinmeyen')} - {param.get('Metot', '')}")
        
        print(f"✅ {migrated_count} parametre migrate edildi")
    except Exception as e:
        print(f"❌ Parametre migrasyon hatası: {e}")

def migrate_firma_olcumler(db, firma_olcum_data):
    """Firma ölçümlerini migrate et"""
    if not firma_olcum_data:
        print("⚠️ Firma ölçüm verisi bulunamadı")
        return
    
    firma_olcum_ref = db.collection('firma_olcumler')
    migrated_count = 0
    
    try:
        # Firma ölçümleri liste formatında
        if isinstance(firma_olcum_data, list):
            for olcum in firma_olcum_data:
                # JSON formatındaki alanları Firestore formatına çevir
                firestore_data = {
                    'firma_adi': olcum.get('firma_adi', ''),
                    'olcum_kodu': olcum.get('olcum_kodu', ''),
                    'baslangic_tarihi': olcum.get('baslangic_tarihi', ''),
                    'bitis_tarihi': olcum.get('bitis_tarihi', ''),
                    'il': olcum.get('il', ''),
                    'ilce': olcum.get('ilce', ''),
                    'yetkili': olcum.get('yetkili', ''),
                    'telefon': olcum.get('telefon', ''),
                    'baca_sayisi': int(olcum.get('baca_sayisi', 1)) if olcum.get('baca_sayisi') else 1,
                    'parametreler': olcum.get('parametreler', []),
                    'baca_parametreleri': olcum.get('baca_parametreleri', {}),
                    'personel': olcum.get('personel', []),
                    'notlar': olcum.get('notlar', ''),
                    'durum': olcum.get('durum', 'Aktif'),
                    'olusturma_tarihi': olcum.get('olusturma_tarihi', ''),
                    'created_at': firestore.SERVER_TIMESTAMP
                }
                
                # ID varsa kullan, yoksa otomatik oluştur
                if olcum.get('id'):
                    firma_olcum_ref.document(olcum['id']).set(firestore_data)
                else:
                    firma_olcum_ref.add(firestore_data)
                
                migrated_count += 1
                print(f"  📝 Firma ölçümü eklendi: {olcum.get('firma_adi', 'Bilinmeyen')}")
        
        print(f"✅ {migrated_count} firma ölçümü migrate edildi")
    except Exception as e:
        print(f"❌ Firma ölçüm migrasyon hatası: {e}")

def migrate_measurements(db, measurements_data):
    """Ölçümleri migrate et"""
    if not measurements_data:
        print("⚠️ Ölçüm verisi bulunamadı")
        return
    
    measurements_ref = db.collection('measurements')
    migrated_count = 0
    
    try:
        # Ölçümler liste formatında
        if isinstance(measurements_data, list):
            for measurement in measurements_data:
                # JSON formatındaki alanları Firestore formatına çevir
                firestore_data = {
                    'firma_adi': measurement.get('firma_adi', ''),
                    'olcum_kodu': measurement.get('olcum_kodu', ''),
                    'olcum_tarihi': measurement.get('olcum_tarihi', ''),
                    'parametre': measurement.get('parametre', ''),
                    'sonuc': measurement.get('sonuc', ''),
                    'birim': measurement.get('birim', ''),
                    'baca_adi': measurement.get('baca_adi', ''),
                    'notlar': measurement.get('notlar', ''),
                    'created_at': firestore.SERVER_TIMESTAMP
                }
                
                # ID varsa kullan, yoksa otomatik oluştur
                if measurement.get('id'):
                    measurements_ref.document(measurement['id']).set(firestore_data)
                else:
                    measurements_ref.add(firestore_data)
                
                migrated_count += 1
                print(f"  📝 Ölçüm eklendi: {measurement.get('firma_adi', 'Bilinmeyen')} - {measurement.get('parametre', '')}")
        
        print(f"✅ {migrated_count} ölçüm migrate edildi")
    except Exception as e:
        print(f"❌ Ölçüm migrasyon hatası: {e}")

def main():
    """Ana migrasyon fonksiyonu"""
    print("🚀 Firebase Veri Migrasyonu Başlatılıyor...")
    print("=" * 50)
    
    # Firebase Admin SDK'yı başlat
    try:
        cred = credentials.Certificate('serviceAccountKey.json')
        firebase_admin.initialize_app(cred)
        db = firestore.client()
        print("✅ Firebase Admin SDK başlatıldı")
    except Exception as e:
        print(f"❌ Firebase başlatma hatası: {e}")
        return
    
    # JSON dosyalarını yükle
    print("\n📂 JSON dosyaları yükleniyor...")
    users_data = load_json_file('users.json')
    parameters_data = load_json_file('parameters.json')
    firma_olcum_data = load_json_file('firma_olcum.json')
    measurements_data = load_json_file('measurements.json')
    
    # Migrasyon işlemlerini başlat
    print("\n🔄 Veri migrasyonu başlatılıyor...")
    
    # Kullanıcıları migrate et
    print("\n👥 Kullanıcılar migrate ediliyor...")
    migrate_users(db, users_data)
    
    # Parametreleri migrate et
    print("\n⚙️ Parametreler migrate ediliyor...")
    migrate_parameters(db, parameters_data)
    
    # Firma ölçümlerini migrate et
    print("\n🏢 Firma ölçümleri migrate ediliyor...")
    migrate_firma_olcumler(db, firma_olcum_data)
    
    # Ölçümleri migrate et
    print("\n📊 Ölçümler migrate ediliyor...")
    migrate_measurements(db, measurements_data)
    
    print("\n" + "=" * 50)
    print("🎉 Migrasyon tamamlandı!")
    print(f"⏰ Tamamlanma zamanı: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("🌐 Uygulama adresi: https://akare-emisyon-saha.web.app")

if __name__ == "__main__":
    main() 
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Otomatik Veri Yedekleme Sistemi
"""

import json
import os
import shutil
from datetime import datetime
import schedule
import time

def create_backup():
    """Tüm veri dosyalarını yedekler"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_dir = f'backups/backup_{timestamp}'
    
    # Yedekleme dizini oluştur
    os.makedirs(backup_dir, exist_ok=True)
    
    # Yedeklenecek dosyalar
    data_files = [
        'baca_bilgileri.json',
        'parametre_olcum.json',
        'firma_olcum.json',
        'users.json',
        'parameters.json'
    ]
    
    for file in data_files:
        if os.path.exists(file):
            shutil.copy2(file, backup_dir)
            print(f"✅ {file} yedeklendi")
    
    # Eski yedekleri temizle (7 günden eski)
    cleanup_old_backups()
    
    print(f"📦 Yedekleme tamamlandı: {backup_dir}")

def cleanup_old_backups():
    """7 günden eski yedekleri temizler"""
    backup_dir = 'backups'
    if not os.path.exists(backup_dir):
        return
    
    current_time = datetime.now()
    for item in os.listdir(backup_dir):
        item_path = os.path.join(backup_dir, item)
        if os.path.isdir(item_path):
            # Dizin oluşturma zamanını kontrol et
            creation_time = datetime.fromtimestamp(os.path.getctime(item_path))
            if (current_time - creation_time).days > 7:
                shutil.rmtree(item_path)
                print(f"🗑️ Eski yedek silindi: {item}")

def main():
    """Ana fonksiyon"""
    print("🔄 Otomatik Yedekleme Sistemi Başlatılıyor...")
    
    # İlk yedekleme
    create_backup()
    
    # Her saat başı yedekleme
    schedule.every().hour.do(create_backup)
    
    # Her gün gece yarısı yedekleme
    schedule.every().day.at("00:00").do(create_backup)
    
    print("⏰ Yedekleme zamanlaması ayarlandı:")
    print("  - Her saat başı")
    print("  - Her gün gece yarısı")
    print("  - 7 günden eski yedekler otomatik silinir")
    
    # Sürekli çalış
    while True:
        schedule.run_pending()
        time.sleep(60)  # 1 dakika bekle

if __name__ == "__main__":
    main()

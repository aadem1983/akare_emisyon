const admin = require('firebase-admin');
const serviceAccount = require('./serviceAccountKey.json');

// Firebase Admin SDK'yı başlat
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

// Personel verileri
const personelVerileri = [
    {
        ad: 'Adem',
        soyad: 'Yıldırım',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif',
        password: '1'
    },
    {
        ad: 'Atalay',
        soyad: 'Toker',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif',
        password: '1'
    },
    {
        ad: 'Kamuran',
        soyad: 'Osman',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif',
        password: '1'
    },
    {
        ad: 'Semih',
        soyad: '',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif',
        password: '1'
    },
    {
        ad: 'Kaan',
        soyad: 'Süzer',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif',
        password: '1'
    },
    {
        ad: 'Emin',
        soyad: '',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif',
        password: '1'
    },
    {
        ad: 'Resul',
        soyad: 'Poyraz',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif',
        password: '1'
    },
    {
        ad: 'Hafize',
        soyad: 'Fazlı',
        gorev: 'Rap',
        rol: '3',
        durum: 'Aktif',
        password: '1'
    }
];

// Personel verilerini Firebase'e yükle
async function uploadPersonel() {
    try {
        console.log('Personel verileri Firebase\'e yükleniyor...');
        
        for (const personel of personelVerileri) {
            await db.collection('personel').add(personel);
            console.log(`${personel.ad} ${personel.soyad} yüklendi`);
        }
        
        console.log('Tüm personel verileri başarıyla yüklendi!');
        process.exit(0);
    } catch (error) {
        console.error('Hata:', error);
        process.exit(1);
    }
}

// Scripti çalıştır
uploadPersonel(); 
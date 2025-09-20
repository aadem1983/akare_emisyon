// Firebase'e personel verilerini güncelleme scripti
const { initializeApp } = require('firebase/app');
const { getFirestore, collection, doc, setDoc, getDocs } = require('firebase/firestore');

// Firebase yapılandırması
const firebaseConfig = {
  apiKey: "AIzaSyC6w4su0eF75ew9zmBVauplAhQ2tAx10WA",
  authDomain: "akare-emisyon-saha.firebaseapp.com",
  projectId: "akare-emisyon-saha",
  storageBucket: "akare-emisyon-saha.firebasestorage.app",
  messagingSenderId: "556818354539",
  appId: "1:556818354539:web:cec60bcb7697d908213c2e"
};

// Firebase'i başlat
const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// Güncellenmiş personel verileri
const personelVerileri = [
    {
        ad: 'Adem',
        soyad: 'Yıldırım',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif'
    },
    {
        ad: 'Atalay',
        soyad: 'Toker',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif'
    },
    {
        ad: 'Kamuran',
        soyad: 'Osman',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif'
    },
    {
        ad: 'Semih',
        soyad: '',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif'
    },
    {
        ad: 'Kaan',
        soyad: 'Süzer',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif'
    },
    {
        ad: 'Emin',
        soyad: '',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif'
    },
    {
        ad: 'Resul',
        soyad: 'Poyraz',
        gorev: 'Saha',
        rol: '3',
        durum: 'Aktif'
    },
    {
        ad: 'Hafize',
        soyad: 'Fazlı',
        gorev: 'Rap',
        rol: '3',
        durum: 'Aktif'
    }
];

// Personel verilerini Firebase'e güncelle
async function updatePersonel() {
  try {
    console.log('Personel verileri Firebase\'e güncelleniyor...');
    
    for (const personel of personelVerileri) {
      // Her personel için benzersiz ID kullan
      const docRef = doc(db, 'personel', personel.ad);
      await setDoc(docRef, personel);
      console.log(`${personel.ad} ${personel.soyad} güncellendi (Görev: ${personel.gorev})`);
    }
    
    console.log('Tüm personel verileri başarıyla güncellendi!');
  } catch (error) {
    console.error('Hata:', error);
  } finally {
    process.exit(0);
  }
}

updatePersonel(); 
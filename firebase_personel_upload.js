const admin = require('firebase-admin');
const fs = require('fs');

// Firebase Admin SDK yapılandırması
const serviceAccount = require('./akare-emisyon-saha-firebase-adminsdk.json');

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

// Personel verilerini oku
const usersData = JSON.parse(fs.readFileSync('./users.json', 'utf8'));

// Personel verilerini Firebase'e yükle
async function uploadPersonel() {
  try {
    console.log('Personel verileri Firebase\'e yükleniyor...');
    
    for (const [username, userData] of Object.entries(usersData)) {
      if (username === 'admin') continue; // Admin kullanıcısını atla
      
      const personelData = {
        ad: username,
        soyad: userData.surname || '',
        gorev: userData.gorev || 'Saha', // Varsayılan olarak Saha
        rol: userData.role || '3',
        durum: 'Aktif',
        created_at: admin.firestore.FieldValue.serverTimestamp()
      };
      
      console.log(`${username} yükleniyor:`, personelData);
      
      await db.collection('personel').doc(username).set(personelData);
    }
    
    console.log('Tüm personel verileri başarıyla yüklendi!');
  } catch (error) {
    console.error('Hata:', error);
  } finally {
    process.exit(0);
  }
}

uploadPersonel(); 
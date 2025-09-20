// Firma ölçüm düzenleme JavaScript'i
function getJsonFromScriptTag(id) {
    const el = document.getElementById(id);
    if (!el) return null;
    try {
        return JSON.parse(el.textContent);
    } catch (e) {
        return el.textContent;
    }
}

window.rawParameters = getJsonFromScriptTag('parameters-json') || [];
window.currentIlce = getJsonFromScriptTag('current-ilce-json') || '';
window.mevcutBacaParametreleri = getJsonFromScriptTag('baca-parametreleri-json') || {};

document.addEventListener('DOMContentLoaded', function() {
    // Parametre listesini global değişkenden al
    var rawParameters = window.rawParameters || [];
    var parametreler = [];
    for (var i = 0; i < rawParameters.length; i++) {
        var param = rawParameters[i];
        parametreler.push({
            id: param.id || '',
            ad: param['Parametre Adı'] || param.ad || param.isim || '',
            metot: param.Metot || param.metot || ''
        });
    }

    // İl-İlçe bağlantısı
    const ilSelect = document.getElementById('il');
    const ilceSelect = document.getElementById('ilce');
    
    // Mevcut ilçeyi yükle
    if (ilSelect.value) {
        loadIlceler(ilSelect.value, window.currentIlce || '');
    }
    
    ilSelect.addEventListener('change', function() {
        const selectedIl = this.value;
        ilceSelect.innerHTML = '<option value="">İlçe Seçiniz</option>';
        
        if (selectedIl) {
            loadIlceler(selectedIl);
        }
    });
    
    function loadIlceler(ilAdi, seciliIlce = '') {
        fetch(`/api/ilceler/${encodeURIComponent(ilAdi)}`)
            .then(response => response.json())
            .then(ilceler => {
                ilceler.forEach(ilce => {
                    const option = document.createElement('option');
                    option.value = ilce;
                    option.textContent = ilce;
                    if (ilce === seciliIlce) {
                        option.selected = true;
                    }
                    ilceSelect.appendChild(option);
                });
            })
            .catch(error => {
                console.error('İlçe verileri yüklenemedi:', error);
            });
    }
    
    // Baca sayısı ve tablo işlemleri
    const bacaSayisiInput = document.getElementById('baca_sayisi');
    const bacaEkleBtn = document.getElementById('bacaEkleBtn');
    const bacaTabloContainer = document.getElementById('bacaTabloContainer');
    const bacaTablo = document.getElementById('bacaTablo');
    
    // Mevcut baca sayısını kontrol et ve tabloyu göster
    const mevcutBacaSayisi = parseInt(bacaSayisiInput.value) || 1;
    if (mevcutBacaSayisi > 0) {
        renderBacaParametreMatrisi(window.mevcutBacaParametreleri || {});
        bacaTabloContainer.style.display = 'block';
    }
    
    // Baca Ekle butonuna tıklandığında
    bacaEkleBtn.addEventListener('click', function() {
        const yeniBacaSayisi = parseInt(bacaSayisiInput.value);
        if (yeniBacaSayisi < 1 || yeniBacaSayisi > 20) {
            alert('Baca sayısı 1-20 arasında olmalıdır!');
            return;
        }
        
        // Mevcut baca sayısını al
        const tbody = bacaTablo.querySelector('tbody');
        const mevcutSatirSayisi = tbody.rows.length;
        
        if (yeniBacaSayisi > mevcutSatirSayisi) {
            // Yeni baca satırları ekle
            yeniBacaSatirlariEkle(yeniBacaSayisi - mevcutSatirSayisi);
        } else if (yeniBacaSayisi < mevcutSatirSayisi) {
            // Fazla satırları sil
            fazlaSatirlariSil(mevcutSatirSayisi - yeniBacaSayisi);
        }
        
        // Baca sayısı input'unu güncelle
        bacaSayisiInput.value = yeniBacaSayisi;
        bacaTabloContainer.style.display = 'block';
    });
    
    // Baca-Parametre Matrisini veriye göre oluşturan fonksiyon
    function renderBacaParametreMatrisi(bacaParametreleriData) {
        // Tablo başlığını temizle ve yeniden oluştur
        const thead = bacaTablo.querySelector('thead');
        thead.innerHTML = `
            <tr>
                <th style="width: 300px;">Baca İsimleri</th>
            </tr>
            <tr>
                <th style="width: 300px;"></th>
            </tr>
        `;
        
        // Tekrarlayan parametreleri filtrele ve benzersiz parametreleri al
        const benzersizParametreler = [];
        const parametreAdlari = new Set();
        
        parametreler.forEach(param => {
            if (!parametreAdlari.has(param.ad)) {
                parametreAdlari.add(param.ad);
                benzersizParametreler.push(param);
            }
        });
        
        // Parametre başlıklarını ekle (sayı göstergesi ile)
        benzersizParametreler.forEach(param => {
            const th = document.createElement('th');
            th.innerHTML = `
                <div class="text-center">
                    <div class="fw-bold">${param.ad}</div>
                    <div class="text-muted small" id="sayac_${param.ad.replace(/\s+/g, '_')}">0</div>
                </div>
            `;
            th.style.minWidth = '60px';
            th.style.maxWidth = '80px';
            th.style.fontSize = '11px';
            th.className = 'text-center';
            thead.querySelector('tr:last-child').appendChild(th);
        });
        
        // Tablo gövdesini temizle ve yeniden oluştur
        const tbody = bacaTablo.querySelector('tbody');
        tbody.innerHTML = '';
        
        // Veri satırlarını oluştur
        Object.entries(bacaParametreleriData).forEach(([bacaAdi, selectedParams], index) => {
            const tr = document.createElement('tr');
            
            // Baca adı hücresi
            const bacaAdiTd = document.createElement('td');
            bacaAdiTd.innerHTML = '<input type="text" class="form-control form-control-sm" ' +
                       'name="baca_adi_' + (index + 1) + '" value="' + bacaAdi + '" ' +
                       'placeholder="Baca ' + (index + 1) + ' ismini yazın..." ' +
                       'style="border: none; background: transparent; box-shadow: none;">';
            tr.appendChild(bacaAdiTd);
            
            // Parametre checkbox'ları
            benzersizParametreler.forEach(function(param) {
                const td = document.createElement('td');
                td.className = 'text-center';
                
                const isChecked = selectedParams.includes(param.ad);
                
                td.innerHTML = '<div class="form-check d-flex justify-content-center">' +
                    '<input class="form-check-input parametre-checkbox" type="checkbox" ' +
                    'name="parametre_' + (index + 1) + '_' + param.ad.replace(/\s+/g, '_') + '" ' +
                    'value="' + param.ad + '" ' +
                    'id="param_' + (index + 1) + '_' + param.ad.replace(/\s+/g, '_') + '" ' +
                    'data-parametre="' + param.ad + '" ' +
                    (isChecked ? 'checked' : '') + '>' +
                    '</div>';
                tr.appendChild(td);
            });
            
            tbody.appendChild(tr);
        });
        
        // Checkbox event listener'larını ekle
        document.querySelectorAll('.parametre-checkbox').forEach(checkbox => {
            checkbox.addEventListener('change', function() {
                sayacGuncelle();
            });
        });
        
        // Sayaçları güncelle
        sayacGuncelle();
    }
    
    // Yeni baca satırları ekleme fonksiyonu
    function yeniBacaSatirlariEkle(eklenecekSayi) {
        const tbody = bacaTablo.querySelector('tbody');
        const mevcutSatirSayisi = tbody.rows.length;
        
        // Parametre listesini al
        const benzersizParametreler = [];
        const parametreAdlari = new Set();
        
        parametreler.forEach(param => {
            if (!parametreAdlari.has(param.ad)) {
                parametreAdlari.add(param.ad);
                benzersizParametreler.push(param);
            }
        });
        
        // Yeni satırları ekle
        for (let i = 0; i < eklenecekSayi; i++) {
            const yeniSatirNo = mevcutSatirSayisi + i + 1;
            const tr = document.createElement('tr');
            
            // Baca adı hücresi
            const bacaAdiTd = document.createElement('td');
            bacaAdiTd.innerHTML = '<input type="text" class="form-control form-control-sm" ' +
                       'name="baca_adi_' + yeniSatirNo + '" value="Baca ' + yeniSatirNo + '" ' +
                       'placeholder="Baca ' + yeniSatirNo + ' ismini yazın..." ' +
                       'style="border: none; background: transparent; box-shadow: none;">';
            tr.appendChild(bacaAdiTd);
            
            // Parametre checkbox'ları
            benzersizParametreler.forEach(function(param) {
                const td = document.createElement('td');
                td.className = 'text-center';
                
                td.innerHTML = '<div class="form-check d-flex justify-content-center">' +
                    '<input class="form-check-input parametre-checkbox" type="checkbox" ' +
                    'name="parametre_' + yeniSatirNo + '_' + param.ad.replace(/\s+/g, '_') + '" ' +
                    'value="' + param.ad + '" ' +
                    'id="param_' + yeniSatirNo + '_' + param.ad.replace(/\s+/g, '_') + '" ' +
                    'data-parametre="' + param.ad + '">' +
                    '</div>';
                tr.appendChild(td);
            });
            
            tbody.appendChild(tr);
        }
        
        // Yeni checkbox'lar için event listener ekle
        document.querySelectorAll('.parametre-checkbox').forEach(checkbox => {
            if (!checkbox.hasEventListener) {
                checkbox.addEventListener('change', function() {
                    sayacGuncelle();
                });
                checkbox.hasEventListener = true;
            }
        });
        
        // Sayaçları güncelle
        sayacGuncelle();
    }
    
    // Fazla satırları silme fonksiyonu
    function fazlaSatirlariSil(silinecekSayi) {
        const tbody = bacaTablo.querySelector('tbody');
        
        // Son satırlardan başlayarak sil
        for (let i = 0; i < silinecekSayi; i++) {
            if (tbody.rows.length > 0) {
                tbody.deleteRow(tbody.rows.length - 1);
            }
        }
        
        // Sayaçları güncelle
        sayacGuncelle();
    }
    
    // Sayaç güncelleme fonksiyonu
    function sayacGuncelle() {
        const benzersizParametreler = [];
        const parametreAdlari = new Set();
        
        parametreler.forEach(param => {
            if (!parametreAdlari.has(param.ad)) {
                parametreAdlari.add(param.ad);
                benzersizParametreler.push(param);
            }
        });
        
        benzersizParametreler.forEach(param => {
            const parametreAdi = param.ad;
            const sayacElement = document.getElementById(`sayac_${parametreAdi.replace(/\s+/g, '_')}`);
            if (sayacElement) {
                const seciliSayisi = document.querySelectorAll(`input[data-parametre="${parametreAdi}"]:checked`).length;
                sayacElement.textContent = seciliSayisi;
            }
        });
    }
    
    // Form gönderildiğinde
    const form = document.getElementById('editForm');
    form.addEventListener('submit', function(e) {
        e.preventDefault();
        
        // Baca parametrelerini topla
        const bacaParametreleri = {};
        const bacaAdiInputs = document.querySelectorAll('input[name^="baca_adi_"]');
        
        bacaAdiInputs.forEach((input) => {
            const bacaAdi = input.value.trim();
            if (bacaAdi) { // Boş baca adlarını atla
                const parametreler = [];
                
                // Input'un name attribute'undan baca numarasını çıkar
                const nameMatch = input.name.match(/baca_adi_(\d+)/);
                if (nameMatch) {
                    const bacaNo = nameMatch[1];
                    
                    // Bu baca için seçili parametreleri bul
                    const parametreCheckboxes = document.querySelectorAll(`input[name^="parametre_${bacaNo}_"]:checked`);
                    parametreCheckboxes.forEach(checkbox => {
                        parametreler.push(checkbox.value);
                    });
                    
                    bacaParametreleri[bacaAdi] = parametreler;
                }
            }
        });
        
        // FormData oluştur
        const formData = new FormData(form);
        formData.append('baca_parametreleri', JSON.stringify(bacaParametreleri));
        
        // Loading göster
        const submitBtn = this.querySelector('button[type="submit"]');
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Güncelleniyor...';
        
        // Formu gönder
        fetch(window.location.href, {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (response.redirected) {
                window.location.href = response.url;
            } else {
                return response.text();
            }
        })
        .then(html => {
            if (html) {
                // Hata durumunda sayfayı yenile
                document.documentElement.innerHTML = html;
            }
        })
        .catch(error => {
            console.error('Hata:', error);
            alert('Bir hata oluştu!');
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i class="fas fa-save"></i> Güncelle';
        });
    });
}); 
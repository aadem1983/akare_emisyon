import pandas as pd

# Test verisi
data = {
    "Sıra": [1, 2, 3],
    "Tarih": ["11.08.25", "15.08.25", "16.08.25"],
    "Firma": ["ATALAY TOKER", "ATALAY TOKER", "Test Firma"],
    "Kod": ["EE-20811-Q2", "DODD", "TEST 001"],
    "Baca": ["SSS", "SSS", "Test Baca"],
    "Personel": ["Admin", "Admin", "Admin"],
    "Cihaz": ["Test Cihaz", "Test Cihaz", "Test Cihaz"],
    "Değer": [510, 480, 500.5]
}

df = pd.DataFrame(data)
df.to_excel("test_rapor.xlsx", index=False, engine='openpyxl')
print("Excel dosyası oluşturuldu: test_rapor.xlsx")


ÇANAKKALE - İSTANBUL GEZİSİ BAŞVURU SİSTEMİ
==========================================

Bu proje tek klasörlük Flask uygulamasıdır.
Excel şablonundan öğrenci yükleyebilir ve veli başvurularını toplayabilirsiniz.

KURULUM
-------
1) Komut satırını bu klasörde açın.
2) Aşağıdaki komutları çalıştırın:

   pip install -r requirements.txt
   python app.py

3) Tarayıcıda şu adresi açın:
   http://127.0.0.1:5000

ADMİN GİRİŞİ
------------
Adres: http://127.0.0.1:5000/admin
Şifre: 1234

ÖNEMLİ
------
Gerçek kullanım öncesi mutlaka değiştirin:
- app.py içindeki SECRET_KEY
- app.py içindeki ADMIN_PASSWORD

EXCEL ŞABLONU
-------------
Varsayılan dosya adı:
8. SINIF ÖĞRENCİLERİ.xlsx

Beklenen sayfa:
Sayfa2

Beklenen sütunlar:
- SINIF
- ÖĞRENCİ NO
- ÖĞRENCİ ADI SOYADI
- TC. KİMLİK NO
- ÖĞRENCİ GRUBU

ÖZELLİKLER
----------
- Ana sayfada kalan süre + boş kontenjan + başvuru durumu
- TC + okul no ile öğrenci doğrulama
- Tek aktif başvuru kontrolü
- Başvuru numarası üretme
- Başvuru durumu sorgulama
- Başvuru iptali
- Admin paneli
- Excel ile öğrenci yükleme
- Ödeme / ek süre tanımlama
- CSV rapor alma
- İşlem logları
- SQLite veritabanı

DOSYALAR
--------
- app.py
- requirements.txt
- README.txt
- 8. SINIF ÖĞRENCİLERİ.xlsx
- gezi.db (ilk çalıştırmada oluşur)

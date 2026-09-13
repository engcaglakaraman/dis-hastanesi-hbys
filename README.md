# Diş Hastanesi Bilgi Yönetim Sistemi

Flask ve Microsoft SQL Server kullanılarak geliştirilmiş, hastane operasyonlarını, hekim randevularını, stok takibini ve finansal raporlamayı tek bir panelde yönetmeyi sağlayan web tabanlı bir bilgi yönetim sistemidir.

## Özellikler

* **Hasta Yönetimi:** Hasta kayıt, güncelleme, tıbbi uyarı/alerji takibi ve detaylı geçmiş arşivi.
* **Hekim ve Randevu Paneli:** Poliklinik hekim kadrosu, branş yönetimi ve günlük randevu takibi.
* **Finans & Kasa Raporları:** Dönemsel ciro takibi (günlük, haftalık, aylık), tahsilat makbuzu yazdırma ve **Excel formatında (.xlsx)** rapor indirme.
* **Z-Raporu (Gün Sonu):** Vezne ve kasa kapatma, ödeme yöntemlerine göre dağılım ve yazdırılabilir günlük özet.
* **Stok & Envanter Takip:** Kritik eşik uyarıları, hızlı miktar güncelleme ve sarf malzeme yönetimi.
* **Modern Arayüz:** Bootstrap 5 ve Bootstrap Icons kullanılarak tasarlanmış, koyu/açık tema (Dark/Light mode) desteğine sahip kullanıcı dostu arayüz.

## Teknolojiler

* **Backend:** Python, Flask, PyODBC
* **Veritabanı:** Microsoft SQL Server (MSSQL)
* **Raporlama:** OpenPyXL (Excel entegrasyonu)
* **Frontend:** Bootstrap 5, Chart.js, HTML5/CSS3

## Kurulum

1. Depoyu klonlayın:
```bash
git clone https://github.com/kullaniciadi/dis-hastanesi-sistemi.git
cd dis-hastanesi-sistemi

```


2. Gerekli kütüphaneleri yükleyin:
```bash
pip install flask pyodbc openpyxl

```


3. SQL Server üzerinde `DISHASTANESI` veritabanını oluşturun ve gerekli tabloları tanımlayın.
4. Uygulamayı çalıştırın:
```bash
python app.py

```


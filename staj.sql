USE DISHASTANESI;
GO

-- =========================================
-- 1. BRANŞLAR
-- =========================================
IF OBJECT_ID(N'dbo.Branslar', N'U') IS NULL
BEGIN
    CREATE TABLE Branslar (
        brans_id INT IDENTITY(1,1) PRIMARY KEY,
        brans_adi NVARCHAR(100) NOT NULL UNIQUE
    );
END
GO

-- =========================================
-- 2. DİŞ HEKİMLERİ
-- =========================================
IF OBJECT_ID(N'dbo.DisHekimleri', N'U') IS NULL
BEGIN
    CREATE TABLE DisHekimleri (
        hekim_id INT IDENTITY(1,1) PRIMARY KEY,
        brans_id INT NOT NULL,
        ad NVARCHAR(50) NOT NULL,
        soyad NVARCHAR(50) NOT NULL,
        telefon NVARCHAR(20),
        email NVARCHAR(100) UNIQUE,
        oda_no NVARCHAR(10),

        CONSTRAINT FK_DisHekimleri_Branslar
            FOREIGN KEY (brans_id)
            REFERENCES Branslar(brans_id)
    );
END
GO

-- =========================================
-- 3. HASTALAR
-- =========================================
IF OBJECT_ID(N'dbo.Hastalar', N'U') IS NULL
BEGIN
    CREATE TABLE Hastalar (
        hasta_id INT IDENTITY(1,1) PRIMARY KEY,
        tc_no CHAR(11) UNIQUE,
        ad NVARCHAR(50) NOT NULL,
        soyad NVARCHAR(50) NOT NULL,
        dogum_tarihi DATE ,
        cinsiyet NVARCHAR(20),
        telefon NVARCHAR(20),
        email NVARCHAR(100) UNIQUE,
        adres NVARCHAR(255),

        CONSTRAINT CK_Hastalar_Cinsiyet
            CHECK (cinsiyet IN (N'Kadın', N'Erkek', N'Belirtilmemiş'))
    );
END
GO

-- =========================================
-- 4. SİGORTALAR
-- =========================================
IF OBJECT_ID(N'dbo.Sigortalar', N'U') IS NULL
BEGIN
    CREATE TABLE Sigortalar (
        sigorta_id INT IDENTITY(1,1) PRIMARY KEY,
        sigorta_adi NVARCHAR(100) NOT NULL UNIQUE
    );
END
GO

-- =========================================
-- 5. HASTA - SİGORTA
-- =========================================
IF OBJECT_ID(N'dbo.HastaSigorta', N'U') IS NULL
BEGIN
    CREATE TABLE HastaSigorta (
        hasta_id INT NOT NULL,
        sigorta_id INT NOT NULL,
        police_no NVARCHAR(50) NOT NULL,
        baslangic_tarihi DATE,
        bitis_tarihi DATE,

        PRIMARY KEY (hasta_id, sigorta_id),

        CONSTRAINT FK_HastaSigorta_Hasta
            FOREIGN KEY (hasta_id)
            REFERENCES Hastalar(hasta_id),

        CONSTRAINT FK_HastaSigorta_Sigorta
            FOREIGN KEY (sigorta_id)
            REFERENCES Sigortalar(sigorta_id),

        CONSTRAINT UQ_HastaSigorta_Police
            UNIQUE (police_no)
    );
END
GO

-- =========================================
-- 6. RANDEVULAR
-- =========================================
IF OBJECT_ID(N'dbo.Randevular', N'U') IS NULL
BEGIN
    CREATE TABLE Randevular (
        randevu_id INT IDENTITY(1,1) PRIMARY KEY,
        hasta_id INT NOT NULL,
        hekim_id INT NOT NULL,
        randevu_tarihi DATE NOT NULL,
        randevu_saati TIME NOT NULL,
        durum NVARCHAR(20) NOT NULL DEFAULT N'Bekliyor',
        notlar NVARCHAR(500),

        CONSTRAINT FK_Randevular_Hasta
            FOREIGN KEY (hasta_id)
            REFERENCES Hastalar(hasta_id),

        CONSTRAINT FK_Randevular_Hekim
            FOREIGN KEY (hekim_id)
            REFERENCES DisHekimleri(hekim_id),

        CONSTRAINT CK_Randevular_Durum
            CHECK (durum IN (N'Bekliyor', N'Tamamlandı', N'İptal', N'Gelmedi')),

        CONSTRAINT UQ_Randevu_Hekim_Tarih_Saat
            UNIQUE (hekim_id, randevu_tarihi, randevu_saati)
    );
END
GO

-- =========================================
-- 7. DİŞ İŞLEMLERİ
-- =========================================
IF OBJECT_ID(N'dbo.DisIslemleri', N'U') IS NULL
BEGIN
    CREATE TABLE DisIslemleri (
        islem_id INT IDENTITY(1,1) PRIMARY KEY,
        islem_adi NVARCHAR(100) NOT NULL UNIQUE,
        aciklama NVARCHAR(500),
        ucret DECIMAL(10,2) NOT NULL,

        CONSTRAINT CK_DisIslemleri_Ucret
            CHECK (ucret >= 0)
    );
END
GO

-- =========================================
-- 8. TEDAVİLER
-- =========================================
IF OBJECT_ID(N'dbo.Tedaviler', N'U') IS NULL
BEGIN
    CREATE TABLE Tedaviler (
        tedavi_id INT IDENTITY(1,1) PRIMARY KEY,
        randevu_id INT NOT NULL,
        tani NVARCHAR(500),
        notlar NVARCHAR(1000),
        baslangic_tarihi DATE NOT NULL,
        bitis_tarihi DATE,

        CONSTRAINT FK_Tedaviler_Randevu
            FOREIGN KEY (randevu_id)
            REFERENCES Randevular(randevu_id),

        CONSTRAINT CK_Tedavi_Tarih
            CHECK (bitis_tarihi IS NULL OR bitis_tarihi >= baslangic_tarihi)
    );
END
GO

-- =========================================
-- 9. TEDAVİ - İŞLEM
-- =========================================
IF OBJECT_ID(N'dbo.TedaviIslemleri', N'U') IS NULL
BEGIN
    CREATE TABLE TedaviIslemleri (
        tedavi_id INT NOT NULL,
        islem_id INT NOT NULL,
        dis_no NVARCHAR(10),
        adet INT NOT NULL DEFAULT 1,
        birim_fiyat DECIMAL(10,2) NOT NULL,
        uygulama_tarihi DATE NOT NULL,

        PRIMARY KEY (tedavi_id, islem_id, dis_no),

        CONSTRAINT FK_TedaviIslemleri_Tedavi
            FOREIGN KEY (tedavi_id)
            REFERENCES Tedaviler(tedavi_id),

        CONSTRAINT FK_TedaviIslemleri_Islem
            FOREIGN KEY (islem_id)
            REFERENCES DisIslemleri(islem_id),

        CONSTRAINT CK_TedaviIslemleri_Adet
            CHECK (adet > 0),

        CONSTRAINT CK_TedaviIslemleri_Fiyat
            CHECK (birim_fiyat >= 0)
    );
END
GO

-- =========================================
-- 10. TAHLİL TÜRLERİ
-- =========================================
IF OBJECT_ID(N'dbo.TahlilTurleri', N'U') IS NULL
BEGIN
    CREATE TABLE TahlilTurleri (
        tahlil_turu_id INT IDENTITY(1,1) PRIMARY KEY,
        tahlil_adi NVARCHAR(100) NOT NULL UNIQUE,
        referans_araligi NVARCHAR(100)
    );
END
GO

-- =========================================
-- 11. TAHLİLLER
-- =========================================
IF OBJECT_ID(N'dbo.Tahliller', N'U') IS NULL
BEGIN
    CREATE TABLE Tahliller (
        tahlil_id INT IDENTITY(1,1) PRIMARY KEY,
        hasta_id INT NOT NULL,
        hekim_id INT NOT NULL,
        tahlil_turu_id INT NOT NULL,
        tahlil_tarihi DATETIME NOT NULL,
        sonuc NVARCHAR(1000),
        aciklama NVARCHAR(500),

        CONSTRAINT FK_Tahliller_Hasta
            FOREIGN KEY (hasta_id)
            REFERENCES Hastalar(hasta_id),

        CONSTRAINT FK_Tahliller_Hekim
            FOREIGN KEY (hekim_id)
            REFERENCES DisHekimleri(hekim_id),

        CONSTRAINT FK_Tahliller_TahlilTuru
            FOREIGN KEY (tahlil_turu_id)
            REFERENCES TahlilTurleri(tahlil_turu_id)
    );
END
GO

-- =========================================
-- 12. ÖDEMELER
-- =========================================
IF OBJECT_ID(N'dbo.Odemeler', N'U') IS NULL
BEGIN
    CREATE TABLE Odemeler (
        odeme_id INT IDENTITY(1,1) PRIMARY KEY,
        hasta_id INT NOT NULL,
        tedavi_id INT NULL,
        odeme_tarihi DATETIME NOT NULL,
        tutar DECIMAL(10,2) NOT NULL,
        odeme_yontemi NVARCHAR(30) NOT NULL,
        durum NVARCHAR(20) NOT NULL DEFAULT N'Ödendi',
        aciklama NVARCHAR(500),

        CONSTRAINT FK_Odemeler_Hasta
            FOREIGN KEY (hasta_id)
            REFERENCES Hastalar(hasta_id),

        CONSTRAINT FK_Odemeler_Tedavi
            FOREIGN KEY (tedavi_id)
            REFERENCES Tedaviler(tedavi_id),

        CONSTRAINT CK_Odemeler_Tutar
            CHECK (tutar > 0),

        CONSTRAINT CK_Odemeler_Yontem
            CHECK (odeme_yontemi IN (N'Nakit', N'Kredi Kartı', N'Banka Kartı', N'Havale')),

        CONSTRAINT CK_Odemeler_Durum
            CHECK (durum IN (N'Ödendi', N'Bekliyor', N'İade'))
    );
END
GO


-- =========================================
-- VERİ EKLEME (INSERT KONTROLLÜ)
-- =========================================

-- 1. BRANŞLAR
INSERT INTO Branslar (brans_adi)
SELECT v.brans_adi
FROM (VALUES 
    (N'Ortodonti'),
    (N'Endodonti'),
    (N'Ağız, Diş ve Çene Cerrahisi'),
    (N'Periodontoloji'),
    (N'Pedodonti (Çocuk Diş Hekimliği)')
) AS v(brans_adi)
WHERE NOT EXISTS (
    SELECT 1 FROM Branslar b WHERE b.brans_adi = v.brans_adi
);

-- 2. DİŞ HEKİMLERİ
INSERT INTO DisHekimleri (brans_id, ad, soyad, telefon, email, oda_no)
SELECT v.brans_id, v.ad, v.soyad, v.telefon, v.email, v.oda_no
FROM (VALUES 
    (1, N'Ali', N'Kaya', '05051112233', 'ali.kaya@dishastanesi.com', '101'),
    (2, N'Merve', N'Şahin', '05052223344', 'merve.sahin@dishastanesi.com', '102'),
    (3, N'Ahmet', N'Demir', '05053334455', 'ahmet.demir@dishastanesi.com', '103'),
    (4, N'Selin', N'Yılmaz', '05054445566', 'selin.yilmaz@dishastanesi.com', '104')
) AS v(brans_id, ad, soyad, telefon, email, oda_no)
WHERE NOT EXISTS (
    SELECT 1 FROM DisHekimleri h WHERE h.email = v.email
);

-- 3. HASTALAR
INSERT INTO Hastalar (tc_no, ad, soyad, dogum_tarihi, cinsiyet, telefon, email, adres)
SELECT v.tc_no, v.ad, v.soyad, v.dogum_tarihi, v.cinsiyet, v.telefon, v.email, v.adres
FROM (VALUES 
    ('12345678901', N'Mehmet', N'Öztürk', '1990-05-15', N'Erkek', '05321112233', 'mehmet.ozturk@gmail.com', N'Kadıköy, İstanbul'),
    ('23456789012', N'Ayşe', N'Karan', '1985-11-20', N'Kadın', '05332223344', 'ayse.karan@gmail.com', N'Çankaya, Ankara'),
    ('34567890123', N'Burak', N'Can', '2001-03-08', N'Erkek', '05443334455', 'burak.can@gmail.com', N'Nilüfer, Bursa'),
    ('45678901234', N'Zeynep', N'Arslan', '1998-09-12', N'Kadın', '05554445566', 'zeynep.arslan@gmail.com', N'Karşıyaka, İzmir')
) AS v(tc_no, ad, soyad, dogum_tarihi, cinsiyet, telefon, email, adres)
WHERE NOT EXISTS (
    SELECT 1 FROM Hastalar hs WHERE hs.tc_no = v.tc_no OR hs.email = v.email
);

-- 4. SİGORTALAR
INSERT INTO Sigortalar (sigorta_adi)
SELECT v.sigorta_adi
FROM (VALUES 
    (N'SGK'),
    (N'Allianz Sigorta'),
    (N'Acıbadem Sigorta'),
    (N'Axa Sigorta')
) AS v(sigorta_adi)
WHERE NOT EXISTS (
    SELECT 1 FROM Sigortalar s WHERE s.sigorta_adi = v.sigorta_adi
);

-- 5. HASTA - SİGORTA
INSERT INTO HastaSigorta (hasta_id, sigorta_id, police_no, baslangic_tarihi, bitis_tarihi)
SELECT v.hasta_id, v.sigorta_id, v.police_no, v.baslangic_tarihi, v.bitis_tarihi
FROM (VALUES 
    (1, 1, 'SGK-998822', '2024-01-01', '2027-01-01'),
    (2, 2, 'ALZ-445566', '2025-06-01', '2027-06-01'),
    (3, 3, 'ACB-112233', '2026-01-15', '2027-01-15')
) AS v(hasta_id, sigorta_id, police_no, baslangic_tarihi, bitis_tarihi)
WHERE NOT EXISTS (
    SELECT 1 FROM HastaSigorta hsi 
    WHERE (hsi.hasta_id = v.hasta_id AND hsi.sigorta_id = v.sigorta_id)
       OR hsi.police_no = v.police_no
);

-- 6. DİŞ İŞLEMLERİ
INSERT INTO DisIslemleri (islem_adi, aciklama, ucret)
SELECT v.islem_adi, v.aciklama, v.ucret
FROM (VALUES 
    (N'Kompozit Dolgu', N'Işınlı kompozit estetik dolgu', 1500.00),
    (N'Kanal Tedavisi (Tek Kök)', N'Tek köklü dişler için kanal tedavisi', 2800.00),
    (N'Diş Taşı Temizliği (Detertraj)', N'Alt ve üst çene plak/taş temizliği', 1200.00),
    (N'Gömülü 20lik Diş Çekimi', N'Cerrahi gömülü diş operasyonu', 4500.00),
    (N'Zirkonyum Kaplama', N'Estetik zirkonyum porselen kron', 5500.00)
) AS v(islem_adi, aciklama, ucret)
WHERE NOT EXISTS (
    SELECT 1 FROM DisIslemleri di WHERE di.islem_adi = v.islem_adi
);

-- 7. TAHLİL TÜRLERİ
INSERT INTO TahlilTurleri (tahlil_adi, referans_araligi)
SELECT v.tahlil_adi, v.referans_araligi
FROM (VALUES 
    (N'Panoramik Röntgen (OPG)', N'Tüm çene görüntüleme'),
    (N'Periapikal Röntgen', N'Tek diş kök görüntüleme'),
    (N'Hemogram (Kan Sayımı)', N'Standart değerler'),
    (N'Diş Tomografisi (CBCT)', N'3 Boyutlu çene ve kemik analizi')
) AS v(tahlil_adi, referans_araligi)
WHERE NOT EXISTS (
    SELECT 1 FROM TahlilTurleri tt WHERE tt.tahlil_adi = v.tahlil_adi
);

-- 8. RANDEVULAR
INSERT INTO Randevular (hasta_id, hekim_id, randevu_tarihi, randevu_saati, durum, notlar)
SELECT v.hasta_id, v.hekim_id, v.randevu_tarihi, v.randevu_saati, v.durum, v.notlar
FROM (VALUES 
    (1, 2, CAST(GETDATE() AS DATE), CAST('10:00' AS TIME), N'Tamamlandı', N'Sol alt azı dişi şiddetli ağrı.'),
    (2, 1, CAST(GETDATE() AS DATE), CAST('11:30' AS TIME), N'Bekliyor', N'Ortodonti tel kontrolü.'),
    (3, 3, CAST(GETDATE() AS DATE), CAST('14:00' AS TIME), N'Bekliyor', N'20lik diş çekim muayenesi.'),
    (4, 4, DATEADD(DAY, 1, CAST(GETDATE() AS DATE)), CAST('15:30' AS TIME), N'Bekliyor', N'Diş eti kanaması şikayeti.')
) AS v(hasta_id, hekim_id, randevu_tarihi, randevu_saati, durum, notlar)
WHERE NOT EXISTS (
    SELECT 1 FROM Randevular r 
    WHERE r.hekim_id = v.hekim_id 
      AND r.randevu_tarihi = v.randevu_tarihi 
      AND r.randevu_saati = v.randevu_saati
);

-- 9. TEDAVİLER
IF NOT EXISTS (SELECT 1 FROM Tedaviler WHERE randevu_id = 1)
BEGIN
    INSERT INTO Tedaviler (randevu_id, tani, notlar, baslangic_tarihi, bitis_tarihi)
    VALUES (1, N'İleri Derece Pulpa İltihabı (Pulpitis)', N'Kanal tedavisi ve dolgu tamamlandı.', CAST(GETDATE() AS DATE), CAST(GETDATE() AS DATE));
END

-- 10. TEDAVİ İŞLEMLERİ
INSERT INTO TedaviIslemleri (tedavi_id, islem_id, dis_no, adet, birim_fiyat, uygulama_tarihi)
SELECT v.tedavi_id, v.islem_id, v.dis_no, v.adet, v.birim_fiyat, v.uygulama_tarihi
FROM (VALUES 
    (1, 2, '36', 1, 2800.00, CAST(GETDATE() AS DATE)),
    (1, 1, '36', 1, 1500.00, CAST(GETDATE() AS DATE))
) AS v(tedavi_id, islem_id, dis_no, adet, birim_fiyat, uygulama_tarihi)
WHERE NOT EXISTS (
    SELECT 1 FROM TedaviIslemleri ti 
    WHERE ti.tedavi_id = v.tedavi_id 
      AND ti.islem_id = v.islem_id 
      AND ti.dis_no = v.dis_no
);

-- 11. TAHLİLLER
IF NOT EXISTS (SELECT 1 FROM Tahliller WHERE hasta_id = 1 AND hekim_id = 2 AND tahlil_turu_id = 2)
BEGIN
    INSERT INTO Tahliller (hasta_id, hekim_id, tahlil_turu_id, tahlil_tarihi, sonuc, aciklama)
    VALUES (1, 2, 2, GETDATE(), N'Kök ucunda hafif lezyon tespit edildi.', N'Periapikal film çekildi, kanala başlandı.');
END

IF NOT EXISTS (SELECT 1 FROM Tahliller WHERE hasta_id = 3 AND hekim_id = 3 AND tahlil_turu_id = 1)
BEGIN
    INSERT INTO Tahliller (hasta_id, hekim_id, tahlil_turu_id, tahlil_tarihi, sonuc, aciklama)
    VALUES (3, 3, 1, GETDATE(), N'Alt çene 38 ve 48 no tam kemik içi gömülü.', N'Cerrahi çekim planlandı.');
END

-- 12. ÖDEMELER
IF NOT EXISTS (SELECT 1 FROM Odemeler WHERE hasta_id = 1 AND tedavi_id = 1)
BEGIN
    INSERT INTO Odemeler (hasta_id, tedavi_id, odeme_tarihi, tutar, odeme_yontemi, durum, aciklama)
    VALUES (1, 1, GETDATE(), 4300.00, N'Kredi Kartı', N'Ödendi', N'Kanal tedavisi ve dolgu ödemesi tahsil edildi.');
END
GO

USE DISHASTANESI;
SELECT * FROM Branslar;
SELECT * FROM DisHekimleri;
SELECT * FROM Hastalar;
SELECT * FROM Sigortalar;
SELECT * FROM Randevular;


USE DISHASTANESI;
GO

IF OBJECT_ID(N'dbo.Receteler', N'U') IS NULL
BEGIN
    CREATE TABLE Receteler (
        recete_id INT IDENTITY(1,1) PRIMARY KEY,
        hasta_id INT NOT NULL,
        hekim_id INT NOT NULL,
        recete_kodu NVARCHAR(20) NOT NULL UNIQUE,
        ilac_adi NVARCHAR(150) NOT NULL,
        dozaj NVARCHAR(50) NOT NULL,
        kullanim_sikligi NVARCHAR(50) NOT NULL,
        kullanim_talimati NVARCHAR(500),
        olusturma_tarihi DATETIME NOT NULL DEFAULT GETDATE(),

        CONSTRAINT FK_Receteler_Hasta FOREIGN KEY (hasta_id) REFERENCES Hastalar(hasta_id),
        CONSTRAINT FK_Receteler_Hekim FOREIGN KEY (hekim_id) REFERENCES DisHekimleri(hekim_id)
    );
END
GO

USE DISHASTANESI;
GO

IF OBJECT_ID(N'dbo.Stoklar', N'U') IS NULL
BEGIN
    CREATE TABLE Stoklar (
        malzeme_id INT IDENTITY(1,1) PRIMARY KEY,
        malzeme_adi NVARCHAR(150) NOT NULL UNIQUE,
        kategori NVARCHAR(100) NOT NULL,
        miktar INT NOT NULL DEFAULT 0,
        birim NVARCHAR(30) NOT NULL DEFAULT N'Adet',
        kritik_seviye INT NOT NULL DEFAULT 10,
        son_guncelleme DATETIME NOT NULL DEFAULT GETDATE(),

        CONSTRAINT CK_Stok_Miktar CHECK (miktar >= 0),
        CONSTRAINT CK_Stok_Kritik CHECK (kritik_seviye >= 0)
    );

    -- Başlangıç Örnek Sarf Malzemeleri
    INSERT INTO Stoklar (malzeme_adi, kategori, miktar, birim, kritik_seviye) VALUES
    (N'Rulo Pamuk (Klinik Tip)', N'Sarf Malzeme', 85, N'Paket', 20),
    (N'Ortodontik Ark Teli (0.016 Niti)', N'Ortodonti', 14, N'Adet', 15),
    (N'Ortodontik Braket Seti (Metal)', N'Ortodonti', 6, N'Set', 10),
    (N'Kompozit Dolgu Tüpü (A2)', N'Restoratif', 25, N'Tüp', 8),
    (N'Lokal Anestezi Ampulü (Artikain)', N'Anestezi', 120, N'Ampul', 30),
    (N'Endodontik Kanal Eğesi (K-File #25)', N'Endodonti', 8, N'Kutu', 10);
END
GO

USE DISHASTANESI;
GO

-- Hastalar tablosuna kronik rahatsızlık / acil durum notu alanı ekleme
IF NOT EXISTS (
    SELECT 1 FROM sys.columns 
    WHERE object_id = OBJECT_ID(N'dbo.Hastalar') AND name = 'tibbi_not'
)
BEGIN
    ALTER TABLE dbo.Hastalar ADD tibbi_not NVARCHAR(500) NULL;
END
GO

USE DISHASTANESI;
GO

-- 1. Hastalar tablosuna elle girilebilir hasta_no (Protokol No) kolonu ekleme
IF NOT EXISTS (
    SELECT 1 FROM sys.columns 
    WHERE object_id = OBJECT_ID(N'dbo.Hastalar') AND name = 'hasta_no'
)
BEGIN
    ALTER TABLE dbo.Hastalar ADD hasta_no NVARCHAR(30) NULL;
END
GO

-- 2. Mevcut kayıtlar için otomatik dosya numarası tanımlama
UPDATE dbo.Hastalar 
SET hasta_no = 'HST-' + RIGHT('0000' + CAST(hasta_id AS VARCHAR(10)), 4)
WHERE hasta_no IS NULL;
GO


USE DISHASTANESI;
GO

-- 1. TC No ve Doğum Tarihi kolonlarındaki NOT NULL zorunluluğunu kaldır
ALTER TABLE dbo.Hastalar ALTER COLUMN tc_no CHAR(11) NULL;
ALTER TABLE dbo.Hastalar ALTER COLUMN dogum_tarihi DATE NULL;

-- 2. Eğer varsa eski kısıtlamaları kaldırıp NULL değerlere izin veren indeksler oluştur
DECLARE @sql NVARCHAR(MAX) = N'';

SELECT @sql += N'ALTER TABLE dbo.Hastalar DROP CONSTRAINT ' + name + ';'
FROM sys.key_constraints
WHERE parent_object_id = OBJECT_ID(N'dbo.Hastalar') AND type = 'UQ';

IF @sql <> N'' EXEC sp_executesql @sql;
GO

-- 3. Boş (NULL) olmayan değerler için tekillik sağlayan filtreli indeksler
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UQ_Hastalar_TC_NotNull')
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UQ_Hastalar_TC_NotNull
    ON dbo.Hastalar(tc_no)
    WHERE tc_no IS NOT NULL;
END
GO


USE DISHASTANESI;
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'TedaviPlanlari')
BEGIN
    CREATE TABLE dbo.TedaviPlanlari (
        plan_id INT IDENTITY(1,1) PRIMARY KEY,
        hasta_id INT NOT NULL,
        hekim_id INT NOT NULL,
        plan_adi NVARCHAR(150) NOT NULL,
        aciklama NVARCHAR(MAX) NULL,
        toplam_seans INT DEFAULT 1,
        tamamlanan_seans INT DEFAULT 0,
        tahmini_tutar DECIMAL(10,2) DEFAULT 0.00,
        durum NVARCHAR(50) DEFAULT N'Devam Ediyor',
        olusturma_tarihi DATETIME DEFAULT GETDATE(),
        CONSTRAINT FK_TedaviPlan_Hasta FOREIGN KEY (hasta_id) REFERENCES Hastalar(hasta_id) ON DELETE CASCADE,
        CONSTRAINT FK_TedaviPlan_Hekim FOREIGN KEY (hekim_id) REFERENCES DisHekimleri(hekim_id)
    );
END
GO
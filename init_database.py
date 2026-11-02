import os
import sys
import csv
import sqlite3
import hashlib
from datetime import datetime, date

sys.stdout.reconfigure(encoding='utf-8')

DB_PATH = r"c:\isg\isg_muratpasa.db"
CSV_283_PATH = r"c:\isg\data\inspection_reference\2026_Muratpasa_Tum_Okul_ve_Kurumlar_Listesi_283_Kurum.csv"
CSV_CONTACTS_PATH = r"c:\isg\data\inspection_reference\2026_Muratpasa_Sorumlu_Okullar_Isveren_ve_Vekili_Iletisim_Listesi.csv"
RISK_SAMPLE_PATH = r"c:\isg\inspect_risk_sample.py"

def hash_pw(password: str) -> str:
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def create_schema(conn):
    cursor = conn.cursor()

    # 1. Kullanıcılar
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        full_name TEXT NOT NULL,
        title TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('koordinator', 'uzman')),
        email TEXT,
        phone TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 2. Okullar ve Kurumlar (283 Kurum)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS schools (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kurum_kodu TEXT UNIQUE NOT NULL,
        kurum_adi TEXT NOT NULL,
        kurum_statusu TEXT,
        bagli_birim TEXT,
        kurum_turu TEXT,
        oncelik_derecesi INTEGER DEFAULT 3,
        tehlike_sinifi TEXT DEFAULT 'Az Tehlikeli',
        adres TEXT,
        telefon TEXT,
        mudur_ad_soyad TEXT,
        mudur_tel TEXT,
        isveren_vekili_ad_soyad TEXT,
        isveren_vekili_tel TEXT,
        sorumlu_uzman_id INTEGER REFERENCES users(id),
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 3. Okul Güvenliği ve Müfettiş Denetim Metrikleri (16 Tablo Karşılığı)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS inspection_metrics (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        school_id INTEGER UNIQUE REFERENCES schools(id) ON DELETE CASCADE,
        adep_var_mi TEXT DEFAULT 'Var',
        adep_tarihi TEXT,
        ek7_olcek_var_mi TEXT DEFAULT 'Var',
        ek7_tarihi TEXT,
        ozel_guvenlik_sayisi INTEGER DEFAULT 0,
        bahce_gorevlisi_sayisi INTEGER DEFAULT 0,
        sertifikali_bahce_gorevlisi_sayisi INTEGER DEFAULT 0,
        gorev_teblig_durumu TEXT,
        metal_dedektor_egitimli_sayisi INTEGER DEFAULT 0,
        okul_gozlem_raporu_sayisi INTEGER DEFAULT 0,
        kamera_sistemi_var_mi TEXT DEFAULT 'Var',
        kgys_entegre_mi TEXT DEFAULT 'Hayır',
        kamera_30gun_kayit_ve_aci_tam_mi TEXT DEFAULT 'Evet',
        giris_cikis_kapi_durumu TEXT DEFAULT 'Tek Kapıdan Girişi Olan Okul',
        el_dedektoru_var_mi TEXT DEFAULT 'Yok',
        kapi_dedektoru_var_mi TEXT DEFAULT 'Yok',
        turnike_var_mi TEXT DEFAULT 'Yok',
        fiziki_engel_turu TEXT DEFAULT 'Duvar + Tel Çit',
        arac_servis_park_yeri_var_mi TEXT DEFAULT 'Var',
        ozel_arac_bahce_engeli TEXT DEFAULT 'Evet',
        ziyaretci_kayit_sistemi TEXT DEFAULT 'Evet',
        cevre_guvenligi_eksiklik_giderildi_mi TEXT,
        guvenlik_olay_sayisi INTEGER DEFAULT 0,
        ihmal_istismar_sayisi INTEGER DEFAULT 0,
        bildirilen_ihmal_istismar_sayisi INTEGER DEFAULT 0,
        erken_uyari_seviye1 INTEGER DEFAULT 0,
        erken_uyari_seviye2 INTEGER DEFAULT 0,
        erken_uyari_seviye3_cde INTEGER DEFAULT 0,
        devamsizlik_aile_ziyareti_sayisi INTEGER DEFAULT 0,
        acil_ihbar_tanitim_tarihi TEXT,
        egitim_yonetici_toplam INTEGER DEFAULT 4,
        egitim_yonetici_alan INTEGER DEFAULT 4,
        egitim_ogretmen_toplam INTEGER DEFAULT 45,
        egitim_ogretmen_alan INTEGER DEFAULT 42,
        egitim_pdr_toplam INTEGER DEFAULT 2,
        egitim_pdr_alan INTEGER DEFAULT 2,
        egitim_ogrenci_toplam INTEGER DEFAULT 750,
        egitim_ogrenci_alan INTEGER DEFAULT 710,
        egitim_veli_toplam INTEGER DEFAULT 500,
        egitim_veli_alan INTEGER DEFAULT 340,
        egitim_diger_toplam INTEGER DEFAULT 12,
        egitim_diger_alan INTEGER DEFAULT 12,
        guvenlik_ekibi_toplanti_sayisi INTEGER DEFAULT 3,
        guvenlik_ekibi_gundem_sayisi INTEGER DEFAULT 8,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 4. Risk Değerlendirmeleri ve DÖF Takibi
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS risk_assessments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        school_id INTEGER REFERENCES schools(id) ON DELETE CASCADE,
        barkod_no TEXT,
        revizyon_no TEXT,
        kontrol_edilen_birim TEXT,
        tehlike TEXT,
        tehlike_unsurlari TEXT,
        riskler TEXT,
        olasilik TEXT,
        siddet TEXT,
        risk_puani INTEGER,
        sorumlular TEXT,
        alinan_onlemler TEXT,
        dayanak_mevzuat TEXT,
        termin_tarihi TEXT,
        dof_durumu TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # 5. Periyodik Teknik Kontroller (Alarm & Uyarı Sistemi)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS periodic_controls (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        school_id INTEGER REFERENCES schools(id) ON DELETE CASCADE,
        ekipman_turu TEXT NOT NULL,
        son_kontrol_tarihi TEXT NOT NULL,
        gecerlilik_ayi INTEGER DEFAULT 12,
        sonraki_kontrol_tarihi TEXT NOT NULL,
        kontrol_yapan_kurum TEXT,
        rapor_sonucu TEXT DEFAULT 'Kullanıma Uygun',
        alarm_seviyesi TEXT DEFAULT 'Gecerli',
        notlar TEXT
    );
    """)

    # 6. Dijital Arşiv & Evrak Yüklemeleri
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS documents_archive (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        school_id INTEGER REFERENCES schools(id) ON DELETE CASCADE,
        belge_turu TEXT NOT NULL,
        belge_adi TEXT NOT NULL,
        dosya_yolu TEXT,
        yukleyen_user_id INTEGER REFERENCES users(id),
        yukleme_tarihi TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    conn.commit()

def seed_data(conn):
    cursor = conn.cursor()

    # 1. Kullanıcıları Ekle
    users_data = [
        ("ramazan.eryildiz", hash_pw("isg2026!"), "Ramazan ERYILDIZ", "İlçe İSG Büro Koordinatörü - A Sınıfı İSG Uzmanı", "koordinator", "ramazaneryildiz@gmail.com", "0505 000 00 00"),
        ("uzman.ahmet", hash_pw("isg2026!"), "Ahmet YILMAZ", "İş Güvenliği Uzmanı", "uzman", "ahmet.isg@meb.gov.tr", "0532 111 22 33"),
        ("uzman.fatma", hash_pw("isg2026!"), "Fatma KAYA", "İş Güvenliği Uzmanı", "uzman", "fatma.isg@meb.gov.tr", "0533 222 33 44")
    ]
    cursor.executemany("""
    INSERT OR IGNORE INTO users (username, password_hash, full_name, title, role, email, phone)
    VALUES (?, ?, ?, ?, ?, ?, ?);
    """, users_data)
    conn.commit()

    # Kullanıcı ID'lerini alalım
    cursor.execute("SELECT id, username FROM users")
    user_map = {row[1]: row[0] for row in cursor.fetchall()}
    coord_id = user_map.get("ramazan.eryildiz", 1)
    u1_id = user_map.get("uzman.ahmet", 2)
    u2_id = user_map.get("uzman.fatma", 3)

    # 2. İletişim tablosunu oku (Müdür ve Md. Yrd. bilgileri için)
    contacts_map = {}
    if os.path.exists(CSV_CONTACTS_PATH):
        with open(CSV_CONTACTS_PATH, "r", encoding="utf-8", errors="ignore") as f:
            reader = csv.reader(f)
            rows = list(reader)
            for r in rows[1:]:
                if len(r) >= 8:
                    k_kod = r[1].strip()
                    mudur = r[3].strip()
                    m_tel = r[4].strip()
                    vekili = r[6].strip()
                    v_tel = r[7].strip()
                    contacts_map[k_kod] = {
                        "mudur": mudur, "m_tel": m_tel,
                        "vekili": vekili, "v_tel": v_tel
                    }

    # 3. 283 Kurum Listesini İçe Aktar
    inserted_schools = 0
    if os.path.exists(CSV_283_PATH):
        with open(CSV_283_PATH, "r", encoding="utf-8", errors="ignore") as f:
            reader = csv.reader(f)
            rows = list(reader)
            for i, r in enumerate(rows[1:]):
                if len(r) < 6:
                    continue
                k_kodu = r[2].strip()
                k_adi = r[3].strip()
                b_birim = r[4].strip()
                k_turu = r[5].strip()

                if not k_kodu or not k_adi:
                    continue

                # Statü belirle
                status = "Özel" if "Özel" in k_adi or "özel" in r[1].lower() else "Resmî"

                # Tehlike sınıfı belirle
                if any(x in k_adi.upper() for x in ["MTAL", "MESLEK", "TEKNİK", "ÇOK PROGRAMLI"]):
                    tehlike = "Tehlikeli"
                    oncelik = 1
                elif any(x in k_adi.upper() for x in ["FEN LİSESİ", "ANADOLU LİSESİ"]):
                    tehlike = "Az Tehlikeli"
                    oncelik = 2
                else:
                    tehlike = "Az Tehlikeli"
                    oncelik = 3

                # İletişim verisi eşleştir
                contact = contacts_map.get(k_kodu, {})
                mudur_adi = contact.get("mudur", "")
                mudur_tel = contact.get("m_tel", "")
                vekili_adi = contact.get("vekili", "")
                vekili_tel = contact.get("v_tel", "")

                # Uzman ataması (Döngüsel olarak Ramazan Bey ve diğer uzmanlara paylaştır)
                assigned_user_id = coord_id if (i % 3 == 0 or oncelik == 1) else (u1_id if i % 3 == 1 else u2_id)

                cursor.execute("""
                INSERT OR IGNORE INTO schools (
                    kurum_kodu, kurum_adi, kurum_statusu, bagli_birim, kurum_turu,
                    oncelik_derecesi, tehlike_sinifi, mudur_ad_soyad, mudur_tel,
                    isveren_vekili_ad_soyad, isveren_vekili_tel, sorumlu_uzman_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                """, (
                    k_kodu, k_adi, status, b_birim, k_turu,
                    oncelik, tehlike, mudur_adi, mudur_tel,
                    vekili_adi, vekili_tel, assigned_user_id
                ))
                inserted_schools += 1

    conn.commit()
    print(f"[+] Toplam eklenen kurum sayısı: {inserted_schools}")

    # 4. Her Okul İçin Varsayılan Denetim Metriklerini Oluştur
    cursor.execute("SELECT id, kurum_kodu, kurum_adi, oncelik_derecesi FROM schools")
    all_schools = cursor.fetchall()
    
    for s in all_schools:
        sid, kkod, kadi, oncelik = s
        is_high_risk = (oncelik == 1)

        cursor.execute("""
        INSERT OR IGNORE INTO inspection_metrics (
            school_id, adep_var_mi, adep_tarihi, ek7_olcek_var_mi, ek7_tarihi,
            ozel_guvenlik_sayisi, bahce_gorevlisi_sayisi, sertifikali_bahce_gorevlisi_sayisi,
            kamera_sistemi_var_mi, kgys_entegre_mi, giris_cikis_kapi_durumu,
            turnike_var_mi, fiziki_engel_turu, arac_servis_park_yeri_var_mi,
            guvenlik_olay_sayisi, guvenlik_ekibi_toplanti_sayisi
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            sid,
            "Var", "15.09.2026",
            "Var", "01.10.2026",
            1 if is_high_risk else 0,
            2 if is_high_risk else 1,
            1 if is_high_risk else 0,
            "Var",
            "Evet" if is_high_risk else "Hayır",
            "Birden Fazla Kapıdan Girişi Olan Okul" if is_high_risk else "Tek Kapıdan Girişi Olan Okul",
            "Var" if is_high_risk else "Yok",
            "Duvar + Tel Çit",
            "Var",
            1 if is_high_risk else 0,
            3
        ))
    conn.commit()
    print(f"[+] 283 okul için denetim metrikleri eşleştirildi.")

    # 5. Örnek Risk Değerlendirmesi Verilerini Yükle (Muratpaşa MTAL & Diğerleri)
    cursor.execute("SELECT id FROM schools WHERE kurum_kodu = '764955'")
    row = cursor.fetchone()
    if row:
        mtal_id = row[0]
        sample_risks = [
            ("970103BRK000001", "29", "İdari Oda", "Elektrik Kabloları", "DAĞINIK KABLO", "Elektrik çarpması, Yangın, Takılıp Düşme", "3. Orta", "5. Çok Ciddi", 15, "ELEKTRİK ALAN ŞEFİ", "TÜM AÇIK KABLOLAR KANAL İLE KORUNACAK", "25/06/2027", "Süresi Devam Ediyor"),
            ("970103BRK000002", "12", "İdari Oda", "Dolap", "DOLAPLARIN DUVARA SABİTLENMEMESİ", "Devrilme, Ezilme", "2. Düşük", "3. Orta", 6, "AHŞAP TEKNOLOJİSİ BÖLÜM ŞEFİ", "DOLAPLARIN DUVARA SABİTLENMESİ", "31/03/2027", "Süresi Devam Ediyor"),
            ("970103BRK000005", "8", "Yangın Merdiveni", "Acil Durum", "YANGIN MERDİVENİNİN BULUNMAMASI", "Çoklu Ölüm, Ezilme", "1. Çok Düşük", "5. Çok Ciddi", 5, "İL DESTEK HİZMETLERİ BİRİMİ", "YANGIN MERDİVENİ YAPILMASI", "30/06/2027", "Süresi Devam Ediyor"),
            ("970103BRK000007", "9", "Koridor", "Yangın Sistemi", "YANGIN DOLABINDA UYARI LEVHASI VE BAKIM EKSİĞİ", "Acil Müdahale Güçlüğü, Yangın", "4. Yüksek", "5. Çok Ciddi", 20, "İŞVEREN", "YANGIN DOLAP BAKIMLARININ YAPILMASI", "30/06/2026", "SÜRESİ GEÇMİŞ (AÇIK DÖF)"),
            ("970103BRK000008", "11", "Merdivenler", "Merdiven Boşluğu", "GÜVENLİK FİLESİNİN BULUNMAMASI", "Ölüm, Düşme, Yaralanma", "1. Çok Düşük", "5. Çok Ciddi", 5, "İŞVEREN VEKİLİ", "GÜVENLİK FİLESİ YAPILMASI", "31/03/2027", "Süresi Devam Ediyor"),
            ("970103BRK000009", "8", "Çatı Katı", "Çatı Kapağı", "ÇATI ÇIKIŞ KAPAĞININ KİLİTLİ OLMAMASI", "Düşme, Yaralanma", "1. Çok Düşük", "3. Orta", 3, "İŞVEREN VEKİLİ", "KİLİTLİ DEMİR KAPAK YAPILMASI", "04/08/2026", "SÜRESİ GEÇMİŞ (AÇIK DÖF)")
        ]
        for r in sample_risks:
            cursor.execute("""
            INSERT OR IGNORE INTO risk_assessments (
                school_id, barkod_no, revizyon_no, kontrol_edilen_birim, tehlike,
                tehlike_unsurlari, riskler, olasilik, siddet, risk_puani,
                sorumlular, alinan_onlemler, termin_tarihi, dof_durumu
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (mtal_id, *r))
        conn.commit()
        print(f"[+] Muratpaşa MTAL için örnek risk ve açık DÖF kayıtları işlendi.")

    # 6. Örnek Periyodik Kontrol ve Alarmlar Ekle
    periodic_samples = [
        (mtal_id, "Elektrik İç Tesisat & Topraklama", "15.11.2025", 12, "15.11.2026", "TMMOB EMO Yetkili Mühendis", "Kullanıma Uygun", "Yaklasiyor (30 Gun)", "Bir sonraki ölçüme 23 gün kaldı."),
        (mtal_id, "Paratoner (Yıldırımdan Korunma)", "10.08.2025", 12, "10.08.2026", "Akredite Muayene Kuruluşu", "Kusurlu (Topraklama direnci yüksek)", "Suresi Gecmis", "DİKKAT: Paratoner kontrol süresi 74 gün geçmiş!"),
        (mtal_id, "Kalorifer Kazanı / Basınçlı Kap", "20.10.2026", 12, "20.10.2027", "TMMOB MMO", "Kullanıma Uygun", "Gecerli", "Periyodik kontrol yeni tamamlandı."),
        (mtal_id, "Yangın Söndürme Tüpleri ve Tesisatı", "01.03.2026", 12, "01.03.2027", "TSE Onaylı Yetkili Firma", "Kullanıma Uygun", "Gecerli", "6 ay sonra hidrostatik basınç testi var.")
    ]
    for p in periodic_samples:
        cursor.execute("""
        INSERT INTO periodic_controls (
            school_id, ekipman_turu, son_kontrol_tarihi, gecerlilik_ayi,
            sonraki_kontrol_tarihi, kontrol_yapan_kurum, rapor_sonucu,
            alarm_seviyesi, notlar
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, p)
    conn.commit()
    print(f"[+] Örnek teknik periyodik muayene ve alarm kayıtları eklendi.")

def main():
    print("=" * 65)
    print(" İLÇE İSG YÖNETİM SİSTEMİ VERİTABANI OLUŞTURULUYOR")
    print("=" * 65)
    conn = sqlite3.connect(DB_PATH)
    create_schema(conn)
    seed_data(conn)
    conn.close()
    print(f"\n[SUCCESS] Veritabanı başarıyla oluşturuldu: {DB_PATH}")

if __name__ == "__main__":
    main()

import os
import sys
import re
import json
import sqlite3
import openpyxl
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding='utf-8')
DB_PATH = r"c:\isg\isg_muratpasa.db"
EXCEL_PATH = r"C:\Users\User\.workspace-mcp\attachments\2026_Muratpasa_268_Okul_Veri_Toplama_ve_Otomatik_Icmal_Cizelgesi_3f263ed8-7660-46a8-b380-e44c19db791a.xlsx"

def clean_phone(val):
    if not val: return ''
    s = str(val).split('.')[0].strip()
    digits = ''.join(c for c in s if c.isdigit())
    if len(digits) == 10 and digits.startswith('5'):
        return '0' + digits
    elif len(digits) == 11 and digits.startswith('05'):
        return digits
    return s

def to_int_safe(val, default=0):
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return int(val)
    m = re.search(r'\d+', str(val))
    if m:
        return int(m.group(0))
    return default

def determine_kademe(tur, ad):
    combined = f"{tur or ''} {ad or ''}".lower()
    if 'mesleki ve teknik' in combined or 'mtal' in combined or 'meslek' in combined or 'mesem' in combined:
        return 'MTAL'
    elif 'imam hatip' in combined or 'ihl' in combined:
        return 'IMAM_HATIP_LISESI'
    elif 'anadolu lisesi' in combined or 'lise' in combined:
        return 'ANADOLU_LISESI'
    elif 'anaokulu' in combined or 'ana okulu' in combined:
        return 'ANAOKULU'
    elif 'ortaokul' in combined:
        return 'ORTAOKUL'
    elif 'ilkokul' in combined:
        return 'ILKOKUL'
    elif 'rehberlik ve arastirma' in combined or 'ram' in combined:
        return 'RAM'
    elif 'bilsem' in combined or 'bilim ve sanat' in combined:
        return 'BILSEM'
    elif 'halk egitimi' in combined or 'hem' in combined:
        return 'HEM'
    elif 'ozel egitim' in combined:
        return 'OZEL_EGITIM'
    return 'ORTAOKUL'

def main():
    print("=" * 70)
    print(" V2.4.0 ENTEGRASYON DİREKTİFİ — VERİTABANI VE VERİ SENKRONİZASYONU")
    print("=" * 70)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # 1. schools tablosuna yeni sütunlar ekle
    c.execute("PRAGMA table_info(schools)")
    existing_school_cols = [r[1] for r in c.fetchall()]
    new_cols = [
        ("is_active_school", "INTEGER DEFAULT 1"),
        ("kategori", "TEXT DEFAULT 'RESMI'"),
        ("genel_mudurluk", "TEXT"),
        ("kademe", "TEXT"),
        ("kolluk_sorumluluk_bolgesi", "TEXT DEFAULT 'Muratpaşa İlçe Emniyet Müdürlüğü'"),
        ("koordinat", "TEXT"),
    ]
    for col, ctype in new_cols:
        if col not in existing_school_cols:
            c.execute(f"ALTER TABLE schools ADD COLUMN {col} {ctype}")
            print(f"[+] schools tablosuna '{col}' sütunu eklendi.")

    # 2. incidents tablosu (Bölüm 2.3)
    c.execute("""
    CREATE TABLE IF NOT EXISTS incidents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        school_id INTEGER REFERENCES schools(id) ON DELETE CASCADE,
        olay_turu TEXT NOT NULL,
        cde_seviye TEXT NOT NULL CHECK(cde_seviye IN ('SEVIYE_1', 'SEVIYE_2', 'SEVIYE_3_CDE')),
        bildirim_112_saati TIMESTAMP,
        bildiren_kisi TEXT,
        cde_toplanma_termini TIMESTAMP,
        olay_sonrasi_degerlendirme_72h BOOLEAN DEFAULT 0,
        toplanti_72h_tarihi TIMESTAMP,
        kvkk_mahremiyet_onayi BOOLEAN DEFAULT 1,
        aciklama TEXT,
        durum TEXT DEFAULT 'Açık' CHECK(durum IN ('Açık', 'ÇDE Toplandı', 'Kapatıldı')),
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)
    print("[+] incidents (Güvenlik Olayları & ÇDE Modeli) tablosu doğrulandı.")

    # 3. okulum_temiz_audits tablosu (Bölüm 2.4)
    c.execute("""
    CREATE TABLE IF NOT EXISTS okulum_temiz_audits (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        school_id INTEGER REFERENCES schools(id) ON DELETE CASCADE,
        tetkik_tarihi TEXT,
        tetkik_saati TEXT DEFAULT '09:30 - 12:30',
        tetkikciler TEXT DEFAULT 'Ramazan ERYILDIZ & Havva AKSU',
        tetkik_turu TEXT DEFAULT 'Gözetim',
        kriter_sonuclari TEXT,
        dof_sayisi INTEGER DEFAULT 0,
        dof_maddeleri TEXT,
        dof_termini TEXT,
        sonuc TEXT DEFAULT 'Planlandı' CHECK(sonuc IN ('Planlandı', 'BELGE_VERILMESI_UYGUN', 'SARTLI_UYGUN_DOF', 'UYGUN_DEGIL')),
        notlar TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)
    print("[+] okulum_temiz_audits (Okulum Temiz Belgelendirme Modeli) tablosu doğrulandı.")
    conn.commit()

    # 4. Excel dosyasını aç ve 268 Okulu Eşitle
    print("\n[+] 268 Okul E-Tablosu okunuyor...")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    s_unt = wb.worksheets[2] # Untitled (268 schools roster)

    sheet_schools = {}
    for r in list(s_unt.iter_rows(values_only=True))[1:]:
        if r[2]:
            code = str(int(r[2]) if isinstance(r[2], float) else r[2]).strip()
            cat_raw = str(r[1]).strip()
            kategori = 'OZEL' if 'zel' in cat_raw else 'RESMI'
            status = 'Özel' if 'zel' in cat_raw else 'Resmî'
            ad = str(r[3]).strip()
            mudurluk = str(r[4]).strip() if r[4] else ''
            tur = str(r[5]).strip() if r[5] else ''
            kademe = determine_kademe(tur, ad)
            tehlike = 'TEHLIKELI' if kademe == 'MTAL' else 'AZ_TEHLIKELI'

            sheet_schools[code] = {
                'kategori': kategori,
                'status': status,
                'ad': ad,
                'mudurluk': mudurluk,
                'tur': tur,
                'kademe': kademe,
                'tehlike': tehlike
            }

    print(f"[+] E-Tablodaki 268 okul analiz edildi: {len(sheet_schools)} kurum.")

    # Veritabanında güncelle: 268 okul aktif, diğer 15 kurum pasif (is_active_school = 0)
    c.execute("UPDATE schools SET is_active_school = 0 WHERE kurum_kodu NOT IN ({})".format(
        ','.join(f"'{c}'" for c in sheet_schools.keys())
    ))
    # 967642 İlçe MEM -> IDARI, yurtlar -> YURT
    c.execute("UPDATE schools SET kategori = 'IDARI' WHERE kurum_kodu = '967642'")
    c.execute("UPDATE schools SET kategori = 'YURT' WHERE kurum_kodu LIKE '999%'")

    for code, info in sheet_schools.items():
        c.execute("""
            UPDATE schools 
            SET is_active_school = 1,
                kategori = ?,
                kurum_statusu = ?,
                genel_mudurluk = ?,
                bagli_birim = ?,
                kurum_turu = ?,
                kademe = ?,
                tehlike_sinifi = ?,
                kolluk_sorumluluk_bolgesi = 'Muratpaşa İlçe Emniyet Müdürlüğü'
            WHERE kurum_kodu = ?
        """, (
            info['kategori'],
            info['status'],
            info['mudurluk'],
            info['mudurluk'],
            info['tur'],
            info['kademe'],
            info['tehlike'],
            code
        ))

    conn.commit()
    print("[+] 268 okulun kategori, kademe ve tehlike sınıfı güncellendi.")

    # 5. Form Yanıtları 1 (95 Resmî Okulun form verilerini aktar)
    print("\n[+] 95 Okulun Form Yanıtları inspection_metrics ve schools tablolarına aktarılıyor...")
    s_resp = wb.worksheets[1]
    resp_rows = list(s_resp.iter_rows(values_only=True))[1:]

    # En güncel yanıtı al (okul bazında son satır)
    by_code = {}
    for r in resp_rows:
        txt = str(r[2])
        m = re.search(r'\b\d{6}\b', txt)
        if m:
            code = m.group(0)
            if code in sheet_schools and sheet_schools[code]['kategori'] == 'RESMI':
                by_code[code] = r

    print(f"[+] Resmî okullara ait tekil form yanıtı sayısı: {len(by_code)}")

    # Önce tüm resmî okulların son_veri_giris_tarihi alanını temizleyelim ki sadece formu dolduran 95 okulda dolu olsun
    c.execute("UPDATE schools SET son_veri_giris_tarihi = NULL, son_veri_giren_kisi = NULL WHERE kategori = 'RESMI'")

    processed_count = 0
    for code, r in by_code.items():
        c.execute("SELECT id FROM schools WHERE kurum_kodu = ?", (code,))
        sch = c.fetchone()
        if not sch:
            continue
        school_id = sch['id']

        ts = str(r[0]) if r[0] else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        giren = str(r[4]) if r[4] else "Okul Müdürü"
        s1_raw = str(r[5]) if r[5] else ""
        m_p = re.search(r'\d+', s1_raw)
        p_derece = int(m_p.group(0)) if m_p else 3

        adep_var = "Var" if "Var" in str(r[6]) else "Yok"
        adep_tarihi = str(r[7]).split()[0] if r[7] else ""
        ek7_var = "Var" if "Var" in str(r[8]) else "Yok"
        ek7_tarihi = str(r[9]).split()[0] if r[9] else ""

        ozel_guv = to_int_safe(r[10])
        bahce = to_int_safe(r[11])
        sertifikali = to_int_safe(r[12])
        gorev_teblig = str(r[13]) if r[13] else ""
        dedektor_egitimli = to_int_safe(r[14])
        gozlem_raporu = to_int_safe(r[15])

        kamera_var = "Var" if "Var" in str(r[16]) else "Yok"
        kgys_entegre = "Evet" if "Evet" in str(r[17]) else "Hayır"
        kamera_30gun = "Evet" if "Evet" in str(r[18]) else "Hayır"

        kapi_durumu = str(r[19]) if r[19] else "Tek Kapı"
        el_ded = "Var" if "Var" in str(r[20]) else "Yok"
        kapi_ded = "Var" if "Var" in str(r[21]) else "Yok"
        turnike = "Var" if "Var" in str(r[22]) else "Yok"
        fiziki_engel = str(r[23]) if r[23] else "Duvar"

        park_yeri = "Var" if "Var" in str(r[24]) else "Yok"
        ozel_arac_engel = "Evet" if "Evet" in str(r[25]) else "Hayır"
        ziyaretci_kayit = "Evet" if "Evet" in str(r[26]) else "Hayır"
        cevre_eksik_giderildi = "Evet" if "Evet" in str(r[27]) else "Hayır"

        olay_sayisi = to_int_safe(r[28])
        ihmal_istismar = to_int_safe(r[29])
        erken_uyari = str(r[30]) if r[30] else ""

        # S26 parsing (Seviye 1, 2, 3)
        nums_s26 = re.findall(r'\d+', erken_uyari)
        sev1 = int(nums_s26[0]) if len(nums_s26) > 0 else 0
        sev2 = int(nums_s26[1]) if len(nums_s26) > 1 else 0
        sev3 = int(nums_s26[2]) if len(nums_s26) > 2 else 0

        # S26-Ek (Aile ziyareti)
        aile_ziyareti_txt = str(r[31]) if r[31] else ""
        nums_az = re.findall(r'\d+', aile_ziyareti_txt)
        aile_ziyareti = int(nums_az[0]) if nums_az else 0

        # S28 (Güvenlik ekibi toplantı)
        ekip_txt = str(r[33]) if r[33] else ""
        nums_ekip = re.findall(r'\d+', ekip_txt)
        toplanti_sayisi = int(nums_ekip[0]) if len(nums_ekip) > 0 else 1
        gundem_sayisi = int(nums_ekip[1]) if len(nums_ekip) > 1 else 3

        # schools tablosunu güncelle
        c.execute("""
            UPDATE schools
            SET oncelik_derecesi = ?,
                son_veri_giris_tarihi = ?,
                son_veri_giren_kisi = ?
            WHERE id = ?
        """, (p_derece, ts, giren, school_id))

        # inspection_metrics tablosunu güncelle / ekle
        c.execute("SELECT id FROM inspection_metrics WHERE school_id = ?", (school_id,))
        m_row = c.fetchone()
        if m_row:
            c.execute("""
                UPDATE inspection_metrics
                SET adep_var_mi = ?, adep_tarihi = ?,
                    ek7_olcek_var_mi = ?, ek7_tarihi = ?,
                    ozel_guvenlik_sayisi = ?, bahce_gorevlisi_sayisi = ?, sertifikali_bahce_gorevlisi_sayisi = ?,
                    gorev_teblig_durumu = ?, metal_dedektor_egitimli_sayisi = ?, okul_gozlem_raporu_sayisi = ?,
                    kamera_sistemi_var_mi = ?, kgys_entegre_mi = ?, kamera_30gun_kayit_ve_aci_tam_mi = ?,
                    giris_cikis_kapi_durumu = ?, el_dedektoru_var_mi = ?, kapi_dedektoru_var_mi = ?, turnike_var_mi = ?,
                    fiziki_engel_turu = ?, arac_servis_park_yeri_var_mi = ?, ozel_arac_bahce_engeli = ?,
                    ziyaretci_kayit_sistemi = ?, cevre_guvenligi_eksiklik_giderildi_mi = ?,
                    guvenlik_olay_sayisi = ?, ihmal_istismar_sayisi = ?,
                    erken_uyari_seviye1 = ?, erken_uyari_seviye2 = ?, erken_uyari_seviye3_cde = ?,
                    devamsizlik_aile_ziyareti_sayisi = ?,
                    guvenlik_ekibi_toplanti_sayisi = ?, guvenlik_ekibi_gundem_sayisi = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE school_id = ?
            """, (
                adep_var, adep_tarihi, ek7_var, ek7_tarihi,
                ozel_guv, bahce, sertifikali,
                gorev_teblig, dedektor_egitimli, gozlem_raporu,
                kamera_var, kgys_entegre, kamera_30gun,
                kapi_durumu, el_ded, kapi_ded, turnike,
                fiziki_engel, park_yeri, ozel_arac_engel,
                ziyaretci_kayit, cevre_eksik_giderildi,
                olay_sayisi, ihmal_istismar,
                sev1, sev2, sev3,
                aile_ziyareti,
                toplanti_sayisi, gundem_sayisi,
                school_id
            ))
        else:
            c.execute("""
                INSERT INTO inspection_metrics (
                    school_id, adep_var_mi, adep_tarihi, ek7_olcek_var_mi, ek7_tarihi,
                    ozel_guvenlik_sayisi, bahce_gorevlisi_sayisi, sertifikali_bahce_gorevlisi_sayisi,
                    gorev_teblig_durumu, metal_dedektor_egitimli_sayisi, okul_gozlem_raporu_sayisi,
                    kamera_sistemi_var_mi, kgys_entegre_mi, kamera_30gun_kayit_ve_aci_tam_mi,
                    giris_cikis_kapi_durumu, el_dedektoru_var_mi, kapi_dedektoru_var_mi, turnike_var_mi,
                    fiziki_engel_turu, arac_servis_park_yeri_var_mi, ozel_arac_bahce_engeli,
                    ziyaretci_kayit_sistemi, cevre_guvenligi_eksiklik_giderildi_mi,
                    guvenlik_olay_sayisi, ihmal_istismar_sayisi,
                    erken_uyari_seviye1, erken_uyari_seviye2, erken_uyari_seviye3_cde,
                    devamsizlik_aile_ziyareti_sayisi,
                    guvenlik_ekibi_toplanti_sayisi, guvenlik_ekibi_gundem_sayisi
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                school_id, adep_var, adep_tarihi, ek7_var, ek7_tarihi,
                ozel_guv, bahce, sertifikali,
                gorev_teblig, dedektor_egitimli, gozlem_raporu,
                kamera_var, kgys_entegre, kamera_30gun,
                kapi_durumu, el_ded, kapi_ded, turnike,
                fiziki_engel, park_yeri, ozel_arac_engel,
                ziyaretci_kayit, cevre_eksik_giderildi,
                olay_sayisi, ihmal_istismar,
                sev1, sev2, sev3,
                aile_ziyareti,
                toplanti_sayisi, gundem_sayisi
            ))

        processed_count += 1

    conn.commit()
    print(f"[+] Toplam {processed_count} resmî okulun verileri güncellendi.")

    # 6. Okulum Temiz Tetkik Programı Tohumlama (Seeding)
    print("\n[+] Okulum Temiz Belgelendirme Tetkik Planı tohumlanıyor...")
    audits_seed = [
        ("702734", "03.11.2026", "09:30 - 12:30", "Ramazan ERYILDIZ & Havva AKSU", "Gözetim", "Planlandı"),
        ("970151", "04.11.2026", "09:30 - 12:30", "Ramazan ERYILDIZ & Havva AKSU", "Gözetim", "Planlandı"),
        ("751351", "05.11.2026", "09:30 - 12:30", "Ramazan ERYILDIZ (Tek Başına)", "Gözetim", "Planlandı"),
        ("969461", "02.11.2026", "09:30 - 12:30", "Bülent AÇIKGÖZ", "Gözetim", "Planlandı"),
        ("969463", "02.11.2026", "09:30 - 12:30", "Havva AKSU & Neşet AKSU", "Gözetim", "Planlandı"),
        ("969457", "02.11.2026", "09:30 - 12:30", "Birsen ERSOY DİNÇ & Metin SAĞLAM", "Gözetim", "Planlandı"),
        ("709661", "03.11.2026", "09:30 - 12:30", "Birsen ERSOY DİNÇ & Metin SAĞLAM", "Gözetim", "Planlandı"),
        ("758804", "04.11.2026", "09:30 - 12:30", "Mete ÖZDEMİR & Metin SAĞLAM", "Gözetim", "Planlandı"),
        ("972199", "02.11.2026", "09:30 - 12:30", "Serdar BASMACI", "Gözetim", "Planlandı"),
        ("709342", "04.11.2026", "09:30 - 12:30", "Serdar BASMACI", "Gözetim", "Planlandı"),
    ]

    c.execute("SELECT count(*) FROM okulum_temiz_audits")
    audit_count = c.fetchone()[0]
    if audit_count == 0:
        for k_kodu, tarih, saat, tetkikci, tur, durum in audits_seed:
            c.execute("SELECT id FROM schools WHERE kurum_kodu = ?", (k_kodu,))
            row = c.fetchone()
            if row:
                s_id = row['id']
                c.execute("""
                    INSERT INTO okulum_temiz_audits (school_id, tetkik_tarihi, tetkik_saati, tetkikciler, tetkik_turu, sonuc)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (s_id, tarih, saat, tetkikci, tur, durum))
        conn.commit()
        print(f"[+] {len(audits_seed)} Okulum Temiz tetkik randevusu oluşturuldu.")
    else:
        print(f"[i] okulum_temiz_audits tablosunda zaten {audit_count} kayıt var.")

    # 7. ÇDE & Güvenlik Olayları Tohumlama (Seeding Initial Incidents with 24h & 48h Countdown)
    print("\n[+] ÇDE ve Güvenlik Olayları Modeli kontrol ediliyor...")
    c.execute("SELECT count(*) FROM incidents")
    inc_count = c.fetchone()[0]
    if inc_count == 0:
        # Örnek 2 kritik Seviye 3 ÇDE vakası ekle (Countdown timer testi için)
        now = datetime.now()
        # 1. Vaka: Ciddi ve yakın tehlike sonrası vaka (Termin: azami 24 saat)
        c.execute("SELECT id FROM schools WHERE kurum_kodu = '970149'") # Antalya MTAL
        s1 = c.fetchone()
        if s1:
            termin_24h = (now + timedelta(hours=18)).strftime("%Y-%m-%d %H:%M:%S")
            c.execute("""
                INSERT INTO incidents (
                    school_id, olay_turu, cde_seviye, bildirim_112_saati, bildiren_kisi,
                    cde_toplanma_termini, olay_sonrasi_degerlendirme_72h, kvkk_mahremiyet_onayi,
                    aciklama, durum
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                s1['id'], 'BICAKLAMA_KESICI', 'SEVIYE_3_CDE', now.strftime("%Y-%m-%d %H:%M:%S"), 'Okul Nöbetçi Öğretmeni',
                termin_24h, 0, 1,
                'Okul bahçe çıkışında kesici aletle yaralanma teşebbüsü. 112 Acil ve Çocuk Şube bilgilendirildi. ÇDE toplanma terminine 18 saat kaldı.',
                'Açık'
            ))

        # 2. Vaka: Yakın tehlike olmayan diğer Seviye 3 vaka (Termin: azami 48 saat)
        c.execute("SELECT id FROM schools WHERE kurum_kodu = '967821'") # Muratpaşa RAM
        s2 = c.fetchone()
        if s2:
            termin_48h = (now + timedelta(hours=41)).strftime("%Y-%m-%d %H:%M:%S")
            c.execute("""
                INSERT INTO incidents (
                    school_id, olay_turu, cde_seviye, bildirim_112_saati, bildiren_kisi,
                    cde_toplanma_termini, olay_sonrasi_degerlendirme_72h, kvkk_mahremiyet_onayi,
                    aciklama, durum
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                s2['id'], 'SIBER_ZORBALIK_MAHREM', 'SEVIYE_3_CDE', now.strftime("%Y-%m-%d %H:%M:%S"), 'Rehberlik Servisi',
                termin_48h, 0, 1,
                'Sosyal medya üzerinden akran şantajı ve mahremiyet ihlali şüphesi. PDR servisi ve RAM eşgüdümüyle ÇDE süreci işletilmektedir.',
                'Açık'
            ))
        conn.commit()
        print("[+] 2 adet ÇDE 24h & 48h vaka kaydı oluşturuldu.")

    # 8. Son Kontroller ve Raporlama
    c.execute("SELECT count(*) FROM schools WHERE is_active_school = 1")
    act_total = c.fetchone()[0]
    c.execute("SELECT count(*) FROM schools WHERE is_active_school = 1 AND kategori = 'RESMI'")
    act_resmi = c.fetchone()[0]
    c.execute("SELECT count(*) FROM schools WHERE is_active_school = 1 AND kategori = 'OZEL'")
    act_ozel = c.fetchone()[0]
    c.execute("SELECT count(*) FROM schools WHERE is_active_school = 1 AND kategori = 'RESMI' AND son_veri_giris_tarihi IS NOT NULL")
    resmi_dolu = c.fetchone()[0]

    print("\n" + "=" * 70)
    print(" VERİTABANI GÜNCELLEMESİ BAŞARIYLA TAMAMLANDI")
    print(f" Aktif Okul Sayısı : {act_total} (Hedef: 268)")
    print(f" Resmî Okul        : {act_resmi} (Hedef: 122)")
    print(f" Özel Okul         : {act_ozel} (Hedef: 146)")
    print(f" Veri Giren Resmî  : {resmi_dolu} (Hedef: 95 - %{(resmi_dolu/act_resmi)*100:.1f})")
    print(f" Kalan Resmî Okul  : {act_resmi - resmi_dolu} (Hedef: 27)")
    print("=" * 70)

    conn.close()

if __name__ == "__main__":
    main()


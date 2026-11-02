import os
import sys
import io
import re
import sqlite3
import unicodedata
import openpyxl
from datetime import datetime
from auth.credential_store import get_credential_store
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding='utf-8')

DB_PATH = r"c:\isg\isg_muratpasa.db"

def norm(s):
    if not s: return ''
    s = str(s).lower().replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ş', 's').replace('ö', 'o').replace('ç', 'c')
    s = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in s if c.isalnum() or c.isspace()).strip()

def clean_phone(val):
    if not val: return ''
    s = str(val).split('.')[0].strip()
    digits = ''.join(c for c in s if c.isdigit())
    if len(digits) == 10 and digits.startswith('5'):
        return '0' + digits
    elif len(digits) == 11 and digits.startswith('05'):
        return digits
def to_int_safe(val, default=0):
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return int(val)
    m = re.search(r'\d+', str(val))
    if m:
        return int(m.group(0))
    return default

def step1_upgrade_schema(conn):
    print("=" * 65)
    print(" 1. VERİTABANI ŞEMASI GENİŞLETİLİYOR")
    print("=" * 65)
    c = conn.cursor()

    # A. schools tablosu
    c.execute("PRAGMA table_info(schools)")
    existing_school_cols = [r[1] for r in c.fetchall()]
    new_school_cols = [
        ("meb_olcek_puani", "REAL"),
        ("kolluk_puani", "REAL"),
        ("toplam_guvenlik_puani", "REAL"),
        ("siginak_var_mi", "TEXT DEFAULT 'Yok'"),
        ("siginak_kapasite_kisi", "INTEGER DEFAULT 0"),
        ("siginak_alan_m2", "INTEGER DEFAULT 0"),
    ]
    for col, ctype in new_school_cols:
        if col not in existing_school_cols:
            c.execute(f"ALTER TABLE schools ADD COLUMN {col} {ctype}")
            print(f"[+] schools tablosuna '{col}' sütunu eklendi.")

    # B. inspection_metrics tablosu
    c.execute("PRAGMA table_info(inspection_metrics)")
    existing_metric_cols = [r[1] for r in c.fetchall()]
    new_metric_cols = [
        ("kamera_bina_ici_calisan", "INTEGER DEFAULT 0"),
        ("kamera_bina_ici_calismayan", "INTEGER DEFAULT 0"),
        ("kamera_bina_disi_calisan", "INTEGER DEFAULT 0"),
        ("kamera_bina_disi_calismayan", "INTEGER DEFAULT 0"),
        ("kamera_kayit_gun_sayisi", "INTEGER DEFAULT 30"),
        ("kamera_ihtiyac_sayisi", "INTEGER DEFAULT 0"),
        ("guvenlik_istihdam_turu", "TEXT"),
        ("guvenlik_izin_belgesi_var_mi", "TEXT DEFAULT 'YOK'"),
        ("sifir_atik_belge_no", "TEXT"),
        ("sifir_atik_belge_tarihi", "TEXT"),
        ("sifir_atik_gecerlilik_tarihi", "TEXT"),
        ("sifir_atik_egitim_sayisi", "INTEGER DEFAULT 0"),
        ("sifir_atik_katilimci_sayisi", "INTEGER DEFAULT 0"),
        ("kazan_durumu", "TEXT DEFAULT 'Var'"),
        ("asansor_durumu", "TEXT DEFAULT 'Yok'"),
        ("su_deposu_durumu", "TEXT DEFAULT 'Var'"),
        ("paratoner_durumu", "TEXT DEFAULT 'Var'"),
    ]
    for col, ctype in new_metric_cols:
        if col not in existing_metric_cols:
            c.execute(f"ALTER TABLE inspection_metrics ADD COLUMN {col} {ctype}")
            print(f"[+] inspection_metrics tablosuna '{col}' sütunu eklendi.")

    # C. school_isg_checklists tablosu
    c.execute("""
    CREATE TABLE IF NOT EXISTS school_isg_checklists (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        school_id INTEGER REFERENCES schools(id) ON DELETE CASCADE,
        kategori_no INTEGER,
        kategori_adi TEXT NOT NULL,
        madde_no TEXT,
        kontrol_kalemi TEXT NOT NULL,
        durum TEXT DEFAULT 'Eksik' CHECK(durum IN ('Var', 'Yok', 'Eksik', 'Muaf')),
        aciklama TEXT,
        denetleyen_uzman_id INTEGER REFERENCES users(id),
        denetim_tarihi TEXT,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(school_id, kategori_adi, kontrol_kalemi)
    );
    """)
    print("[+] school_isg_checklists tablosu doğrulandı.")
    conn.commit()

def step2_seed_checklists(conn):
    print("\n" + "=" * 65)
    print(" 2. 14 KATEGORİ VE 42 İSG KONTROL KALEMİ CHECKLIST ŞABLONU YÜKLENİYOR")
    print("=" * 65)
    checklist_path = r"c:\isg\data\drive_samples\OKUL İSG KLASÖRÜ CHECK LİST mete.xlsx"
    if not os.path.exists(checklist_path):
        print("[-] Checklist dosyası bulunamadı, atlanıyor.")
        return

    wb = openpyxl.load_workbook(checklist_path, data_only=True)
    ws = wb.active

    items = []
    current_kat_no = 1
    current_kat_adi = "Genel"

    for r in range(4, ws.max_row + 1):
        c1 = ws.cell(r, 1).value
        c2 = ws.cell(r, 2).value
        c3 = ws.cell(r, 3).value
        if c1 and str(c1).replace('.', '').strip().isdigit():
            current_kat_no = int(str(c1).replace('.', '').strip())
            if c2:
                current_kat_adi = str(c2).strip()
        elif c2 and not c3:
            current_kat_adi = str(c2).strip()
        elif c3:
            kalem = str(c3).strip()
            kat = str(c2).strip() if c2 else current_kat_adi
            items.append((current_kat_no, kat, str(len(items)+1), kalem))

    print(f"[+] Toplam {len(items)} kontrol kalemi checklist şablonundan okundu.")

    c = conn.cursor()
    c.execute("SELECT id FROM schools")
    school_ids = [r[0] for r in c.fetchall()]

    inserted_count = 0
    for s_id in school_ids:
        for kat_no, kat_adi, m_no, kalem in items:
            c.execute("""
                INSERT OR IGNORE INTO school_isg_checklists 
                (school_id, kategori_no, kategori_adi, madde_no, kontrol_kalemi, durum)
                VALUES (?, ?, ?, ?, ?, 'Eksik')
            """, (s_id, kat_no, kat_adi, m_no, kalem))
            inserted_count += 1
    conn.commit()
    print(f"[OK] {len(school_ids)} okul için toplam {inserted_count} checklist kontrol kalemi bağlandı.")

def step3_import_archive_data(conn):
    print("\n" + "=" * 65)
    print(" 3. 'isg_eski_evraklar' ARŞİV VERİLERİ SİSTEME AKTARILIYOR (Zenginleştirme)")
    print("=" * 65)
    c = conn.cursor()

    # Okul haritası (kurum_kodu -> id ve norm_adi -> id)
    c.execute("SELECT id, kurum_kodu, kurum_adi FROM schools")
    rows = c.fetchall()
    code_map = {r[1]: r[0] for r in rows}
    name_map = {norm(r[2]): (r[0], r[1], r[2]) for r in rows}

    def find_school(code=None, name=None):
        if code:
            c_str = str(code).split('.')[0].strip()
            if c_str in code_map:
                return code_map[c_str]
        if name:
            n_name = norm(name)
            if n_name in name_map:
                return name_map[n_name][0]
            for k, v in name_map.items():
                if n_name in k or k in n_name:
                    return v[0]
        return None

    # A. Antalya Emniyet Bolgesi Derecelendirme.xlsx
    emniyet_path = r"c:\isg\data\drive_samples\Antalya Emniyet Bolgesi Derecelendirme.xlsx"
    if os.path.exists(emniyet_path):
        wb = openpyxl.load_workbook(emniyet_path, data_only=True)
        ws = wb['Sayfa1'] if 'Sayfa1' in wb.sheetnames else wb.active
        upd_emn = 0
        for r in range(2, min(ws.max_row + 1, 350)):
            kk = ws.cell(r, 1).value
            meb_p = ws.cell(r, 3).value
            kol_p = ws.cell(r, 4).value
            top_p = ws.cell(r, 5).value
            derece = ws.cell(r, 6).value
            s_id = find_school(code=kk)
            if s_id:
                try:
                    c.execute("""
                        UPDATE schools 
                        SET meb_olcek_puani = ?, kolluk_puani = ?, toplam_guvenlik_puani = ?,
                            oncelik_derecesi = COALESCE(?, oncelik_derecesi)
                        WHERE id = ?
                    """, (
                        float(meb_p) if meb_p is not None else None,
                        float(kol_p) if kol_p is not None else None,
                        float(top_p) if top_p is not None else None,
                        int(derece) if derece is not None else None,
                        s_id
                    ))
                    upd_emn += 1
                except Exception as e:
                    pass
        conn.commit()
        print(f"[+] Emniyet Derecelendirme Puanları: {upd_emn} okul güncellendi.")

    # B. Kamera_ve_Guvenlik_Gorevlisi.xlsx
    kamera_path = r"c:\isg\data\drive_samples\Kamera_ve_Guvenlik_Gorevlisi.xlsx"
    if os.path.exists(kamera_path):
        wb = openpyxl.load_workbook(kamera_path, data_only=True)
        ws = wb.active
        upd_kam = 0
        for r in range(10, ws.max_row + 1):
            kk = ws.cell(r, 5).value
            s_id = find_school(code=kk)
            if s_id:
                kam_var = ws.cell(r, 6).value
                bina_ici_calisan = ws.cell(r, 7).value or 0
                bina_ici_calismayan = ws.cell(r, 8).value or 0
                bina_disi_calisan = ws.cell(r, 9).value or 0
                bina_disi_calismayan = ws.cell(r, 10).value or 0
                toplam_kam = ws.cell(r, 11).value or 0
                depolama_gun = ws.cell(r, 12).value or 30
                ihtiyac_kam = ws.cell(r, 13).value or 0
                
                ogg_var = ws.cell(r, 15).value
                ogg_sayisi = ws.cell(r, 16).value or 0
                ogg_tur = ws.cell(r, 17).value

                c.execute("""
                    UPDATE inspection_metrics
                    SET kamera_bina_ici_calisan = ?,
                        kamera_bina_ici_calismayan = ?,
                        kamera_bina_disi_calisan = ?,
                        kamera_bina_disi_calismayan = ?,
                        toplam_kamera_sayisi = CASE WHEN toplam_kamera_sayisi IS NULL OR toplam_kamera_sayisi <= 0 THEN ? ELSE toplam_kamera_sayisi END,
                        kamera_kayit_gun_sayisi = ?,
                        kamera_ihtiyac_sayisi = ?,
                        ozel_guvenlik_sayisi = CASE WHEN ozel_guvenlik_sayisi IS NULL OR ozel_guvenlik_sayisi <= 0 THEN ? ELSE ozel_guvenlik_sayisi END,
                        guvenlik_istihdam_turu = ?
                    WHERE school_id = ?
                """, (
                    to_int_safe(bina_ici_calisan), to_int_safe(bina_ici_calismayan),
                    to_int_safe(bina_disi_calisan), to_int_safe(bina_disi_calismayan),
                    to_int_safe(toplam_kam), to_int_safe(depolama_gun, 30), to_int_safe(ihtiyac_kam),
                    to_int_safe(ogg_sayisi), str(ogg_tur).strip() if ogg_tur else None,
                    s_id
                ))
                upd_kam += 1
        conn.commit()
        print(f"[+] Kamera ve Güvenlik Görevlisi: {upd_kam} okul güncellendi.")

    # C. Muratpaşa Sıfır Atık Bilgileri1.xlsx
    sifir_atik_path = r"c:\isg\data\drive_samples\Muratpaşa Sıfır Atık Bilgileri1.xlsx"
    if os.path.exists(sifir_atik_path):
        wb = openpyxl.load_workbook(sifir_atik_path, data_only=True)
        ws = wb['SIFIR ATIK TÜM LİSTE']
        upd_sa = 0
        for r in range(4, ws.max_row + 1):
            kk = ws.cell(r, 3).value
            s_id = find_school(code=kk)
            if s_id:
                b_no = ws.cell(r, 6).value
                b_tar = ws.cell(r, 7).value
                b_gec = ws.cell(r, 8).value
                e_say = ws.cell(r, 9).value or 0
                k_say = ws.cell(r, 10).value or 0

                b_tar_str = b_tar.strftime('%d.%m.%Y') if isinstance(b_tar, datetime) else str(b_tar or '').strip()
                b_gec_str = b_gec.strftime('%d.%m.%Y') if isinstance(b_gec, datetime) else str(b_gec or '').strip()

                c.execute("""
                    UPDATE inspection_metrics
                    SET sifir_atik_belge_no = ?,
                        sifir_atik_belge_tarihi = ?,
                        sifir_atik_gecerlilik_tarihi = ?,
                        sifir_atik_egitim_sayisi = ?,
                        sifir_atik_katilimci_sayisi = ?
                    WHERE school_id = ?
                """, (
                    str(b_no).strip() if b_no else None,
                    b_tar_str if b_tar_str else None,
                    b_gec_str if b_gec_str else None,
                    to_int_safe(e_say),
                    to_int_safe(k_say),
                    s_id
                ))
                upd_sa += 1
        conn.commit()
        print(f"[+] Sıfır Atık Belge ve Eğitim Bilgileri: {upd_sa} okul güncellendi.")

    # D. MURATPASA_Okul_Siginak_Verileri_Tablosu.xlsx
    siginak_path = r"c:\isg\data\drive_samples\MURATPASA_Okul_Siginak_Verileri_Tablosu.xlsx"
    if os.path.exists(siginak_path):
        wb = openpyxl.load_workbook(siginak_path, data_only=True)
        ws = wb.active
        upd_sig = 0
        for r in range(4, ws.max_row + 1):
            okul_adi = ws.cell(r, 6).value
            s_id = find_school(name=okul_adi)
            if s_id:
                kisi = ws.cell(r, 7).value or 0
                alan = ws.cell(r, 8).value or 0
                durum = ws.cell(r, 9).value or 'Sığınak Var'
                try:
                    c.execute("""
                        UPDATE schools
                        SET siginak_var_mi = ?, siginak_kapasite_kisi = ?, siginak_alan_m2 = ?
                        WHERE id = ?
                    """, (
                        'Var' if 'var' in str(durum).lower() else 'Yok',
                        to_int_safe(kisi),
                        to_int_safe(alan),
                        s_id
                    ))
                    upd_sig += 1
                except Exception:
                    pass
        conn.commit()
        print(f"[+] Sığınak Kapasite ve m² Verileri: {upd_sig} okul güncellendi.")

    # E. Muratpaşa Okul ve Kurumların Güvenlik Görevlisi Çalıştırma İzin Belgesi.xlsx
    izin_path = r"c:\isg\data\drive_samples\Muratpaşa Okul ve Kurumların Güvenlik Görevlisi Çalıştırma İzin Belgesi.xlsx"
    if os.path.exists(izin_path):
        wb = openpyxl.load_workbook(izin_path, data_only=True)
        ws = wb.active
        upd_izn = 0
        for r in range(2, ws.max_row + 1):
            okul_adi = ws.cell(r, 1).value
            izin_durum = ws.cell(r, 2).value
            s_id = find_school(name=okul_adi)
            if s_id and izin_durum:
                c.execute("""
                    UPDATE inspection_metrics
                    SET guvenlik_izin_belgesi_var_mi = ?
                    WHERE school_id = ?
                """, ('VAR' if 'var' in str(izin_durum).lower() else 'YOK', s_id))
                upd_izn += 1
        conn.commit()
        print(f"[+] Güvenlik Görevlisi Çalıştırma İzin Belgesi: {upd_izn} okul güncellendi.")

    # F. MURATPAŞA DEDEKTÖRÜ BULUNAN OKUL LİSTESİ.xlsx
    dedektor_path = r"c:\isg\data\drive_samples\MURATPAŞA DEDEKTÖRÜ BULUNAN OKUL LİSTESİ.xlsx"
    if os.path.exists(dedektor_path):
        wb = openpyxl.load_workbook(dedektor_path, data_only=True)
        ws = wb.active
        upd_ded = 0
        for r in range(4, ws.max_row + 1):
            okul_adi = ws.cell(r, 3).value
            s_id = find_school(name=okul_adi)
            if s_id:
                c.execute("""
                    UPDATE inspection_metrics
                    SET el_dedektoru_var_mi = 'Var'
                    WHERE school_id = ?
                """, (s_id,))
                upd_ded += 1
        conn.commit()
        print(f"[+] Dedektörü Bulunan Okullar: {upd_ded} okul güncellendi.")

def step4_apply_ground_truth_from_isg_forms(conn):
    print("\n" + "=" * 65)
    print(" 4. 'İSG' ANA KLASÖRÜNDEN EN GÜNCEL FORM YANITLARI AKTARILIYOR (GROUND TRUTH)")
    print("=" * 65)

    store = get_credential_store()
    creds = store.get_credential('ramazaneryildiz@gmail.com')
    drive = build('drive', 'v3', credentials=creds)

    fid = '1-xYvDp4oaV5TciGeSzizJhF76Sh3bsrtRaVCfSGzT-c'
    print("[+] Drive'dan '2026_Muratpasa_268_Okul_Veri_Toplama_ve_Otomatik_Icmal_Cizelgesi' indiriliyor...")
    req = drive.files().export_media(fileId=fid, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    fh = io.BytesIO(req.execute())
    wb = openpyxl.load_workbook(fh, data_only=True)
    ws = wb['Form Yanıtları 1']

    c = conn.cursor()
    c.execute("SELECT id, kurum_kodu, kurum_adi FROM schools")
    rows = c.fetchall()
    code_map = {r[1]: r[0] for r in rows}
    name_map = {norm(r[2]): r[0] for r in rows}

    upd_count = 0

    for r in range(2, ws.max_row + 1):
        ts = ws.cell(r, 1).value
        email = ws.cell(r, 2).value
        k1_text = str(ws.cell(r, 3).value or '').strip()
        k2_status = ws.cell(r, 4).value
        k3_contact = str(ws.cell(r, 5).value or '').strip()

        # Kurum Kodunu k1_text içinden ayıkla (Örn: "762760 - Hamidiye İmam Hatip...")
        code_match = re.search(r'\b(\d{6,8})\b', k1_text)
        s_id = None
        if code_match:
            kk = code_match.group(1)
            if kk in code_map:
                s_id = code_map[kk]
        if not s_id:
            # İsme göre dene
            n_text = norm(k1_text)
            for k, sid in name_map.items():
                if k in n_text or n_text in k:
                    s_id = sid
                    break

        if not s_id:
            continue

        # Tarih formatı
        ts_str = ts.strftime('%d.%m.%Y %H:%M') if isinstance(ts, datetime) else str(ts or '')

        # İletişim ayrıştırma (Örn: "Hüseyin AKCA- 05372271824")
        contact_name = k3_contact
        contact_phone = clean_phone(k3_contact)
        if '-' in k3_contact:
            parts = k3_contact.split('-')
            contact_name = parts[0].strip()
            if not contact_phone and len(parts) > 1:
                contact_phone = clean_phone(parts[1])

        # Soruların değerlerini al
        s1 = ws.cell(r, 6).value # Öncelik
        s2 = ws.cell(r, 7).value # ADEP Var/Yok
        s3 = ws.cell(r, 8).value # ADEP Tarih
        s4 = ws.cell(r, 9).value # EK-7 Var/Yok
        s5 = ws.cell(r, 10).value # EK-7 Tarih
        s6 = ws.cell(r, 11).value # Özel Güvenlik Sayısı
        s7 = ws.cell(r, 12).value # Bahçe Görevlisi Sayısı
        s8 = ws.cell(r, 13).value # Sertifikalı Bahçe Görevlisi
        s9 = ws.cell(r, 14).value # Tebliğ durumu / tarihi
        s10 = ws.cell(r, 15).value # Metal Dedektör Eğitimi
        s11 = ws.cell(r, 16).value # Okul Gözlem Formu Sayısı
        s12 = ws.cell(r, 17).value # Kamera Sistemi Var/Yok
        s13 = ws.cell(r, 18).value # KGYS Entegre mi
        s14 = ws.cell(r, 19).value # Kamera kayıt yeterli mi
        s15 = ws.cell(r, 20).value # Kapı Durumu
        s16 = ws.cell(r, 21).value # El Dedektörü
        s17 = ws.cell(r, 22).value # Kapı Dedektörü
        s18 = ws.cell(r, 23).value # Turnike
        s19 = ws.cell(r, 24).value # Fiziki Engel
        s20 = ws.cell(r, 25).value # Park Yeri
        s21 = ws.cell(r, 26).value # Özel Araç Bahçe Engeli
        s22 = ws.cell(r, 27).value # Ziyaretçi Kayıt
        s23 = ws.cell(r, 28).value # Çevre Güvenliği Eksiklik Giderildi mi
        s24 = ws.cell(r, 29).value # Olay Sayısı
        s25 = ws.cell(r, 30).value # İhmal/İstismar Sayısı
        s26 = ws.cell(r, 31).value # Erken Uyarı Seviye 1/2/3
        s26_ek = ws.cell(r, 32).value # Aile Ziyareti
        s27 = ws.cell(r, 33).value # Eğitim Sayıları
        s28 = ws.cell(r, 34).value # Yönetim Ekibi Toplantı

        # Öncelik derecesi çözümleme (1, 2, 3, 4)
        oncelik_val = None
        if s1:
            m = re.search(r'(\d)', str(s1))
            if m: oncelik_val = int(m.group(1))

        # Tarih formatlama
        def fmt_d(val):
            if isinstance(val, datetime): return val.strftime('%d.%m.%Y')
            return str(val or '').strip()

        # Sayı ayrıştırma
        def to_i(val, default=0):
            if val is None: return default
            digits = ''.join(c for c in str(val) if c.isdigit())
            return int(digits) if digits else default

        # 1. schools tablosunu güncelle (GÜNCEL GROUND TRUTH)
        c.execute("""
            UPDATE schools
            SET oncelik_derecesi = COALESCE(?, oncelik_derecesi),
                mudur_ad_soyad = CASE WHEN mudur_ad_soyad IS NULL OR mudur_ad_soyad = '' THEN ? ELSE mudur_ad_soyad END,
                mudur_tel = CASE WHEN mudur_tel IS NULL OR mudur_tel = '' THEN ? ELSE mudur_tel END,
                isveren_vekili_ad_soyad = CASE WHEN isveren_vekili_ad_soyad IS NULL OR isveren_vekili_ad_soyad = '' THEN ? ELSE isveren_vekili_ad_soyad END,
                isveren_vekili_tel = CASE WHEN isveren_vekili_tel IS NULL OR isveren_vekili_tel = '' THEN ? ELSE isveren_vekili_tel END,
                eposta = CASE WHEN (eposta IS NULL OR eposta = '') AND ? != '' THEN ? ELSE eposta END,
                son_veri_giris_tarihi = ?,
                son_veri_giren_kisi = ?
            WHERE id = ?
        """, (
            oncelik_val,
            contact_name, contact_phone,
            contact_name, contact_phone,
            str(email or '').strip(), str(email or '').strip(),
            ts_str, contact_name,
            s_id
        ))

        # 2. inspection_metrics tablosunu güncelle (GÜNCEL GROUND TRUTH)
        c.execute("""
            UPDATE inspection_metrics
            SET adep_var_mi = CASE WHEN ? LIKE '%Var%' THEN 'Var' WHEN ? LIKE '%Yok%' THEN 'Yok' ELSE adep_var_mi END,
                adep_tarihi = CASE WHEN ? != '' THEN ? ELSE adep_tarihi END,
                ek7_olcek_var_mi = CASE WHEN ? LIKE '%Var%' THEN 'Var' WHEN ? LIKE '%Yok%' THEN 'Yok' ELSE ek7_olcek_var_mi END,
                ek7_tarihi = CASE WHEN ? != '' THEN ? ELSE ek7_tarihi END,
                ozel_guvenlik_sayisi = ?,
                bahce_gorevlisi_sayisi = ?,
                sertifikali_bahce_gorevlisi_sayisi = ?,
                gorev_teblig_durumu = ?,
                metal_dedektor_egitimli_sayisi = ?,
                okul_gozlem_raporu_sayisi = ?,
                kamera_sistemi_var_mi = CASE WHEN ? LIKE '%Var%' THEN 'Var' WHEN ? LIKE '%Yok%' THEN 'Yok' ELSE kamera_sistemi_var_mi END,
                kgys_entegre_mi = CASE WHEN ? LIKE '%Evet%' THEN 'Evet' WHEN ? LIKE '%Hayır%' THEN 'Hayır' ELSE kgys_entegre_mi END,
                kamera_30gun_kayit_ve_aci_tam_mi = CASE WHEN ? LIKE '%Evet%' THEN 'Evet' WHEN ? LIKE '%Hayır%' THEN 'Hayır' ELSE kamera_30gun_kayit_ve_aci_tam_mi END,
                giris_cikis_kapi_durumu = COALESCE(?, giris_cikis_kapi_durumu),
                el_dedektoru_var_mi = CASE WHEN ? LIKE '%Var%' THEN 'Var' WHEN ? LIKE '%Yok%' THEN 'Yok' ELSE el_dedektoru_var_mi END,
                kapi_dedektoru_var_mi = CASE WHEN ? LIKE '%Var%' THEN 'Var' WHEN ? LIKE '%Yok%' THEN 'Yok' ELSE kapi_dedektoru_var_mi END,
                turnike_var_mi = CASE WHEN ? LIKE '%Var%' THEN 'Var' WHEN ? LIKE '%Yok%' THEN 'Yok' ELSE turnike_var_mi END,
                fiziki_engel_turu = COALESCE(?, fiziki_engel_turu),
                arac_servis_park_yeri_var_mi = CASE WHEN ? LIKE '%Var%' THEN 'Var' WHEN ? LIKE '%Yok%' THEN 'Yok' ELSE arac_servis_park_yeri_var_mi END,
                ozel_arac_bahce_engeli = CASE WHEN ? LIKE '%Evet%' THEN 'Evet' WHEN ? LIKE '%Hayır%' THEN 'Hayır' ELSE ozel_arac_bahce_engeli END,
                ziyaretci_kayit_sistemi = CASE WHEN ? LIKE '%Evet%' THEN 'Evet' WHEN ? LIKE '%Hayır%' THEN 'Hayır' ELSE ziyaretci_kayit_sistemi END,
                cevre_guvenligi_eksiklik_giderildi_mi = COALESCE(?, cevre_guvenligi_eksiklik_giderildi_mi),
                guvenlik_olay_sayisi = ?,
                guvenlik_ekibi_toplanti_sayisi = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE school_id = ?
        """, (
            str(s2 or ''), str(s2 or ''),
            fmt_d(s3), fmt_d(s3),
            str(s4 or ''), str(s4 or ''),
            fmt_d(s5), fmt_d(s5),
            to_i(s6), to_i(s7), to_i(s8),
            str(s9 or '').strip(),
            to_i(s10), to_i(s11),
            str(s12 or ''), str(s12 or ''),
            str(s13 or ''), str(s13 or ''),
            str(s14 or ''), str(s14 or ''),
            str(s15 or '').strip() if s15 else None,
            str(s16 or ''), str(s16 or ''),
            str(s17 or ''), str(s17 or ''),
            str(s18 or ''), str(s18 or ''),
            str(s19 or '').strip() if s19 else None,
            str(s20 or ''), str(s20 or ''),
            str(s21 or ''), str(s21 or ''),
            str(s22 or ''), str(s22 or ''),
            str(s23 or '').strip() if s23 else None,
            to_i(s24), to_i(s28),
            s_id
        ))

        upd_count += 1

    conn.commit()
    print(f"[OK] Toplam {upd_count} okulun verisi en güncel Google Form Yanıtları ile güncellendi ve önceliklendirildi!")

def main():
    conn = sqlite3.connect(DB_PATH)
    try:
        step1_upgrade_schema(conn)
        step2_seed_checklists(conn)
        step3_import_archive_data(conn)
        step4_apply_ground_truth_from_isg_forms(conn)
        print("\n" + "=" * 65)
        print(" TÜM VERİ GENİŞLETME VE ENTEGRASYON İŞLEMLERİ BAŞARIYLA TAMAMLANDI!")
        print("=" * 65)
    finally:
        conn.close()

if __name__ == "__main__":
    main()

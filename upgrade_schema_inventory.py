"""
Envanter Oluşturma Asistanı Veritabanı Şeması ve Drive Master Envanter İçe Aktarımı
Muratpaşa İlçe MEM İSG Portalı - 2026
"""

import os
import re
import sqlite3
import openpyxl

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "isg_muratpasa.db")
DRIVE_XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "2026_Sorumlu_6_MTAL_MESEM_Kurumsal_ISG_ve_Tesis_Envanteri.xlsx")


def classify_space_type(name: str) -> str:
    n = (name or "").lower()
    if any(k in n for k in ["derslik", "sınıf", "şube"]):
        return "Derslik"
    if any(k in n for k in ["atölye", "imalat", "torna", "cnc", "kaynak", "mobilya", "motor", "otomotiv", "tesisat", "boya"]):
        return "Atölye"
    if any(k in n for k in ["lab", "laboratuvar", "uyap", "duruşma"]):
        return "Laboratuvar"
    if any(k in n for k in ["kazan", "ısı merkezi", "kalorifer"]):
        return "Kazan / Isı Merkezi"
    if any(k in n for k in ["trafo", "jeneratör", "elektrik odası", "pano", "sistem odası"]):
        return "Elektrik / Trafo Odası"
    if any(k in n for k in ["müdür", "idari", "öğretmen", "rehberlik", "memur", "arşiv", "fotokopi"]):
        return "İdari Oda"
    if any(k in n for k in ["yemekhane", "mutfak", "kantin", "çay ocağı"]):
        return "Yemekhane / Kantin"
    if any(k in n for k in ["depo", "ambar", "kimyasal"]):
        return "Depo / Arşiv"
    if any(k in n for k in ["sığınak"]):
        return "Sığınak"
    if any(k in n for k in ["spor", "konferans", "kütüphane", "mescid"]):
        return "Sosyal / Spor Alanı"
    if any(k in n for k in ["koridor", "merdiven", "tuvalet", "wc", "lavabo", "giriş", "acil çıkış", "kat"]):
        return "Ortak Alan / Koridor"
    if any(k in n for k in ["bahçe", "çatı", "otopark", "ihata", "sundurma"]):
        return "Bahçe / Çatı"
    return "Genel Mekân"


def create_tables(conn: sqlite3.Connection):
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS school_spaces (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            school_id INTEGER NOT NULL,
            bina_adi TEXT DEFAULT 'Ana Bina',
            kat TEXT DEFAULT 'Zemin Kat',
            mekan_adi TEXT NOT NULL,
            mekan_turu TEXT DEFAULT 'Genel Mekân',
            alan_m2 REAL,
            kapasite INTEGER,
            durum TEXT DEFAULT 'DEVAM_EDIYOR',
            notlar TEXT,
            kaynak TEXT DEFAULT 'UZMAN_GIRISI',
            son_guncelleme TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            guncelleyen_uzman TEXT,
            UNIQUE(school_id, bina_adi, kat, mekan_adi),
            FOREIGN KEY (school_id) REFERENCES schools(id)
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS school_inventory_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            school_id INTEGER NOT NULL,
            space_id INTEGER,
            ekipman_adi TEXT NOT NULL,
            kategori TEXT DEFAULT 'Genel Demirbaş',
            adet INTEGER DEFAULT 1,
            marka_model_seri TEXT,
            uretim_yili TEXT,
            calisma_durumu TEXT DEFAULT 'Aktif',
            periyodik_kontrol_gerekli INTEGER DEFAULT 0,
            son_kontrol_tarihi TEXT,
            kontrol_gecerlilik_ayi INTEGER DEFAULT 12,
            kontrol_etiketi TEXT DEFAULT 'Gerekmiyor',
            isg_guvenlik_durumu TEXT DEFAULT 'Uygun',
            isg_eksik_detay TEXT,
            fotograf_yolu TEXT,
            linked_risk_id INTEGER,
            linked_periodic_id INTEGER,
            kaynak TEXT DEFAULT 'UZMAN_GIRISI',
            ekleyen_uzman TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (school_id) REFERENCES schools(id),
            FOREIGN KEY (space_id) REFERENCES school_spaces(id)
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS school_inventory_meta (
            school_id INTEGER PRIMARY KEY,
            calisan_temsilcisi TEXT,
            sivil_savunma_amiri TEXT,
            resmi_alan_sayisi INTEGER DEFAULT 0,
            dogrulanmis_meslek_alanlari TEXT,
            operasyonel_not TEXT,
            guvenlik_oncelik TEXT,
            ozel_guvenlik_sayisi TEXT,
            bahce_gorevlisi_sayisi TEXT,
            kamera_durumu TEXT,
            kgys_durumu TEXT,
            kayit_suresi_yeterli TEXT,
            el_dedektoru TEXT,
            kapi_dedektoru TEXT,
            turnike TEXT,
            cevre_engeli TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (school_id) REFERENCES schools(id)
        )
    """)

    conn.commit()
    print("[OK] school_spaces, school_inventory_items ve school_inventory_meta tabloları hazır.")


def resolve_school_id(conn: sqlite3.Connection, code_raw: str, name_raw: str):
    c = conn.cursor()
    code_str = str(code_raw or "").replace(".0", "").strip()
    # "969926 / 767197" gibi birleşik kodları ilk koda göre al
    first_code = code_str.split("/")[0].strip()

    if first_code:
        c.execute("SELECT id, kurum_adi FROM schools WHERE kurum_kodu = ?", (first_code,))
        r = c.fetchone()
        if r:
            return r[0], r[1]

    # Muratpaşa MESEM (777376) / Muratpaşa MTAL (764955) eşleştirme
    if "777376" in code_str or "MESEM" in (name_raw or "").upper():
        c.execute("SELECT id, kurum_adi FROM schools WHERE kurum_kodu = '764955' OR kurum_kodu = '777376' ORDER BY id DESC LIMIT 1")
        r = c.fetchone()
        if r:
            # Muratpaşa MTAL (id 83, sorumlu_uzman_id=1) üzerinde tutalım eğer 764955 aktifse
            c.execute("SELECT id, kurum_adi FROM schools WHERE kurum_kodu = '764955'")
            mtal = c.fetchone()
            if mtal:
                return mtal[0], mtal[1]
            return r[0], r[1]

    return None, None


def import_drive_inventory(conn: sqlite3.Connection):
    if not os.path.exists(DRIVE_XLSX):
        print(f"[UYARI] Drive dosyası bulunamadı: {DRIVE_XLSX}")
        return

    wb = openpyxl.load_workbook(DRIVE_XLSX, data_only=True)
    c = conn.cursor()

    # 1. Kurumsal & Yönetim Envanteri + 3. Fiziki Güvenlik Envanteri
    ws1 = wb["1_KURUMSAL_VE_YONETIM_ENVANTERI"]
    ws3 = wb["3_FIZIKI_GUVENLIK_VE_KAMERA_ENV"]

    sec_by_code = {}
    for r in range(2, ws3.max_row):
        code_val = str(ws3.cell(r, 1).value or "").replace(".0", "").strip()
        if not code_val or code_val == "TOPLAM":
            continue
        sec_by_code[code_val] = {
            "guvenlik_oncelik": str(ws3.cell(r, 3).value or "").strip(),
            "ozel_guvenlik_sayisi": str(ws3.cell(r, 4).value or "").replace(".0", "").strip(),
            "bahce_gorevlisi_sayisi": str(ws3.cell(r, 5).value or "").replace(".0", "").strip(),
            "kamera_durumu": str(ws3.cell(r, 7).value or "").strip(),
            "kgys_durumu": str(ws3.cell(r, 8).value or "").strip(),
            "kayit_suresi_yeterli": str(ws3.cell(r, 9).value or "").strip(),
            "el_dedektoru": str(ws3.cell(r, 11).value or "").strip(),
            "kapi_dedektoru": str(ws3.cell(r, 12).value or "").strip(),
            "turnike": str(ws3.cell(r, 13).value or "").strip(),
            "cevre_engeli": str(ws3.cell(r, 14).value or "").strip(),
        }

    imported_schools = []
    for r in range(2, ws1.max_row):
        sira = ws1.cell(r, 1).value
        if str(sira).strip().upper() == "TOPLAM" or not sira:
            continue
        code_val = str(ws1.cell(r, 2).value or "").replace(".0", "").strip()
        name_val = str(ws1.cell(r, 3).value or "").strip()
        school_id, db_name = resolve_school_id(conn, code_val, name_val)
        if not school_id:
            print(f"[ATLANDI] Okul bulunamadı: {code_val} - {name_val}")
            continue

        imported_schools.append((school_id, db_name))
        calisan_tem = str(ws1.cell(r, 11).value or "").strip()
        sivil_sav = str(ws1.cell(r, 12).value or "").strip()
        alan_say = ws1.cell(r, 13).value or 0
        try:
            alan_say_int = int(float(alan_say))
        except Exception:
            alan_say_int = 0
        meslek_alanlari = str(ws1.cell(r, 14).value or "").strip()
        op_not = str(ws1.cell(r, 15).value or "").strip()
        adres = str(ws1.cell(r, 6).value or "").strip()
        sabit_tel = str(ws1.cell(r, 7).value or "").strip()

        # Okul adres ve telefon bilgisini eksikse güncelle
        c.execute("""
            UPDATE schools SET
                adres = COALESCE(NULLIF(adres, ''), ?),
                telefon = COALESCE(NULLIF(telefon, ''), ?)
            WHERE id = ?
        """, (adres, sabit_tel, school_id))

        sec = sec_by_code.get(code_val, {})
        c.execute("""
            INSERT INTO school_inventory_meta (
                school_id, calisan_temsilcisi, sivil_savunma_amiri,
                resmi_alan_sayisi, dogrulanmis_meslek_alanlari, operasyonel_not,
                guvenlik_oncelik, ozel_guvenlik_sayisi, bahce_gorevlisi_sayisi,
                kamera_durumu, kgys_durumu, kayit_suresi_yeterli,
                el_dedektoru, kapi_dedektoru, turnike, cevre_engeli
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(school_id) DO UPDATE SET
                calisan_temsilcisi=excluded.calisan_temsilcisi,
                sivil_savunma_amiri=excluded.sivil_savunma_amiri,
                resmi_alan_sayisi=excluded.resmi_alan_sayisi,
                dogrulanmis_meslek_alanlari=excluded.dogrulanmis_meslek_alanlari,
                operasyonel_not=excluded.operasyonel_not,
                guvenlik_oncelik=excluded.guvenlik_oncelik,
                ozel_guvenlik_sayisi=excluded.ozel_guvenlik_sayisi,
                bahce_gorevlisi_sayisi=excluded.bahce_gorevlisi_sayisi,
                kamera_durumu=excluded.kamera_durumu,
                kgys_durumu=excluded.kgys_durumu,
                kayit_suresi_yeterli=excluded.kayit_suresi_yeterli,
                el_dedektoru=excluded.el_dedektoru,
                kapi_dedektoru=excluded.kapi_dedektoru,
                turnike=excluded.turnike,
                cevre_engeli=excluded.cevre_engeli,
                updated_at=CURRENT_TIMESTAMP
        """, (
            school_id, calisan_tem, sivil_sav, alan_say_int, meslek_alanlari, op_not,
            sec.get("guvenlik_oncelik", ""), sec.get("ozel_guvenlik_sayisi", ""), sec.get("bahce_gorevlisi_sayisi", ""),
            sec.get("kamera_durumu", ""), sec.get("kgys_durumu", ""), sec.get("kayit_suresi_yeterli", ""),
            sec.get("el_dedektoru", ""), sec.get("kapi_dedektoru", ""), sec.get("turnike", ""), sec.get("cevre_engeli", "")
        ))

    # 2. Teknik Tesis & Ekipman Envanterini Mekân ve Ekipman Olarak İşle
    ws2 = wb["2_TEKNIK_TESIS_VE_EKIPMAN_ENVAN"]
    spaces_added = 0
    items_added = 0

    tech_cols = [
        (4, "Basınçlı Kap", 1),
        (5, "Kazan / Isı Merkezi", 1),
        (6, "Kaldırma / İletme", 1),
        (7, "Elektrik / Pano / Trafo", 1),
        (8, "Elektrik / Pano / Trafo", 1),
        (9, "Yangın / Acil Durum", 1),
    ]

    for r in range(2, ws2.max_row):
        code_val = str(ws2.cell(r, 1).value or "").replace(".0", "").strip()
        name_val = str(ws2.cell(r, 2).value or "").strip()
        if not code_val or "GENEL ÖZET" in code_val:
            continue
        school_id, db_name = resolve_school_id(conn, code_val, name_val)
        if not school_id:
            continue

        # 2a. Atölye / Uygulama Birimleri -> school_spaces
        birimler_str = str(ws2.cell(r, 3).value or "").strip()
        if birimler_str:
            for birim in [b.strip() for b in birimler_str.split(",") if b.strip()]:
                # Parantez içi açıklamaları koruyarak temiz ad al
                mtype = classify_space_type(birim)
                c.execute("""
                    INSERT OR IGNORE INTO school_spaces (school_id, bina_adi, kat, mekan_adi, mekan_turu, kaynak, guncelleyen_uzman)
                    VALUES (?, 'Ana Bina', 'Genel', ?, ?, 'DRIVE_ENVANTER', 'Ramazan ERYILDIZ')
                """, (school_id, birim, mtype))
                if c.rowcount > 0:
                    spaces_added += 1

        # 2b. MEBBİS raporlarından gelen mevcut birimleri de mekân olarak tohumla
        c.execute("""
            SELECT DISTINCT kontrol_edilen_birim FROM risk_assessments
            WHERE school_id = ? AND kontrol_edilen_birim IS NOT NULL AND TRIM(kontrol_edilen_birim) != ''
        """, (school_id,))
        for (meb_unit,) in c.fetchall():
            meb_unit = meb_unit.strip()
            mtype = classify_space_type(meb_unit)
            c.execute("""
                INSERT OR IGNORE INTO school_spaces (school_id, bina_adi, kat, mekan_adi, mekan_turu, kaynak, guncelleyen_uzman)
                VALUES (?, 'Ana Bina', 'Genel', ?, ?, 'MEBBIS_BIRIM', 'Ramazan ERYILDIZ')
            """, (school_id, meb_unit, mtype))
            if c.rowcount > 0:
                spaces_added += 1

        # 2c. Teknik ekipman sütunlarını başlangıç envanteri olarak ekle (tekrar etmemek için kontrol et)
        c.execute("SELECT COUNT(*) FROM school_inventory_items WHERE school_id = ? AND kaynak = 'DRIVE_ENVANTER'", (school_id,))
        if c.fetchone()[0] == 0:
            for col_idx, kat_adi, pk_gerekli in tech_cols:
                val = str(ws2.cell(r, col_idx).value or "").strip()
                if not val or val.lower().startswith("yok") or val == "-":
                    continue
                has_issue = any(k in val.lower() for k in ["yok", "eksik", "kilitsiz", "brk", "döf", "askıda"])
                isg_durum = "Eksik / Riskli" if has_issue else "Uygun"
                c.execute("""
                    INSERT INTO school_inventory_items (
                        school_id, space_id, ekipman_adi, kategori, adet,
                        calisma_durumu, periyodik_kontrol_gerekli, kontrol_etiketi,
                        isg_guvenlik_durumu, isg_eksik_detay, kaynak, ekleyen_uzman
                    ) VALUES (?, NULL, ?, ?, 1, 'Aktif', ?, 'Kontrol Edilecek', ?, ?, 'DRIVE_ENVANTER', 'Ramazan ERYILDIZ')
                """, (school_id, val[:220], kat_adi, pk_gerekli, isg_durum, val if has_issue else None))
                items_added += 1

    conn.commit()
    print(f"[OK] Drive master envanterinden {len(imported_schools)} okul meta verisi, {spaces_added} fiziki mekân ve {items_added} teknik ekipman kaydı işlendi.")


if __name__ == "__main__":
    conn = sqlite3.connect(DB_PATH)
    create_tables(conn)
    import_drive_inventory(conn)
    conn.close()

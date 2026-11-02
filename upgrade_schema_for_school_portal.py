import sqlite3
import secrets
import hashlib

DB_PATH = r"c:\isg\isg_muratpasa.db"

def upgrade_schema():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # 1. schools tablosuna yeni sütunlar ekle
    school_cols = [
        ("portal_sifre", "TEXT DEFAULT 'isg2026'"),
        ("access_token", "TEXT"),
        ("ogrenci_sayisi", "INTEGER DEFAULT 0"),
        ("ogretmen_sayisi", "INTEGER DEFAULT 0"),
        ("derslik_sayisi", "INTEGER DEFAULT 0"),
        ("bina_sayisi", "INTEGER DEFAULT 1"),
        ("eposta", "TEXT"),
        ("son_veri_giris_tarihi", "TEXT"),
        ("son_veri_giren_kisi", "TEXT")
    ]

    c.execute("PRAGMA table_info(schools)")
    existing_school_cols = [row[1] for row in c.fetchall()]

    for col_name, col_def in school_cols:
        if col_name not in existing_school_cols:
            c.execute(f"ALTER TABLE schools ADD COLUMN {col_name} {col_def}")
            print(f"[+] schools tablosuna '{col_name}' sütunu eklendi.")

    # 2. inspection_metrics tablosuna yeni sütunlar ekle
    metric_cols = [
        ("isg_kurulu_var_mi", "TEXT DEFAULT 'Evet'"),
        ("calisan_temsilcisi_var_mi", "TEXT DEFAULT 'Evet'"),
        ("ilkyardimci_personel_sayisi", "INTEGER DEFAULT 2"),
        ("tatbikat_tarihi", "TEXT"),
        ("tatbikat_turu", "TEXT DEFAULT 'Yangın ve Tahliye'"),
        ("tatbikat_katilimci_sayisi", "INTEGER DEFAULT 0"),
        ("toplam_kamera_sayisi", "INTEGER DEFAULT 16"),
        ("kor_nokta_kamera_eksigi", "INTEGER DEFAULT 0"),
        ("kazan_kontrol_tarihi", "TEXT"),
        ("elektrik_topraklama_tarihi", "TEXT"),
        ("paratoner_kontrol_tarihi", "TEXT"),
        ("yangin_tupleri_kontrol_tarihi", "TEXT"),
        ("asansor_muayene_tarihi", "TEXT"),
        ("su_deposu_temizlik_tarihi", "TEXT")
    ]

    c.execute("PRAGMA table_info(inspection_metrics)")
    existing_metric_cols = [row[1] for row in c.fetchall()]

    for col_name, col_def in metric_cols:
        if col_name not in existing_metric_cols:
            c.execute(f"ALTER TABLE inspection_metrics ADD COLUMN {col_name} {col_def}")
            print(f"[+] inspection_metrics tablosuna '{col_name}' sütunu eklendi.")

    # 3. Her okul için benzersiz erişim token'ı oluştur (Eğer yoksa)
    c.execute("SELECT id, kurum_kodu FROM schools WHERE access_token IS NULL OR access_token = ''")
    schools_to_token = c.fetchall()
    for s_id, k_kodu in schools_to_token:
        # Kurum koduna dayalı güvenli ama deterministik/rastgele token
        token = hashlib.sha256(f"muratpasa_isg_{k_kodu}_{secrets.token_hex(4)}".encode('utf-8')).hexdigest()[:16]
        c.execute("UPDATE schools SET access_token = ? WHERE id = ?", (token, s_id))

    conn.commit()
    print(f"[+] Toplam {len(schools_to_token)} okul için güvenli erişim token'ı üretildi.")

    conn.close()

if __name__ == "__main__":
    upgrade_schema()

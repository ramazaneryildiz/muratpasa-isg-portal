"""
Muratpaşa İlçe MEM İSG Bilgi Portalı
Veritabanı Şeması v2.6.0 Güncelleme Scripti (upgrade_schema_v26.py)
Google Drive Entegrasyon Direktifi v2.6.0 Uyarınca:
1. feature_requests tablosu oluşturulur (Modül 7: AI Destekli Geliştirici İstek Havuzu).
2. risk_assessments tablosuna Fine-Kinney, MEBBİS köprüleme ve atölye alanları eklenir (Modül 8).
"""

import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "isg_muratpasa.db")

def upgrade_schema():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    print("[*] Veritabanı şeması v2.6.0 güncelleniyor...")

    # 1. feature_requests tablosu
    c.execute("""
        CREATE TABLE IF NOT EXISTS feature_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kurum_kodu TEXT,
            school_id INTEGER,
            gonderen_ad_soyad TEXT NOT NULL,
            unvan TEXT,
            iletisim_bilgisi TEXT,
            istek_turu TEXT DEFAULT 'DIGER',
            talep_basligi TEXT NOT NULL,
            detayli_aciklama TEXT NOT NULL,
            ornek_dokuman_veya_form_yapisi TEXT,
            ai_asistan_ozeti_ve_prompt TEXT,
            ai_oneri_kategori TEXT,
            durum TEXT DEFAULT 'BEKLEMEDE',
            gelistirici_notu TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (school_id) REFERENCES schools(id)
        )
    """)
    print("[+] 'feature_requests' tablosu hazırlandı.")

    # 2. risk_assessments tablosu genişletme
    c.execute("PRAGMA table_info(risk_assessments)")
    existing_cols = [r[1] for r in c.fetchall()]

    new_risk_cols = [
        ("bolum_unite", "TEXT"),
        ("meslek_alani_veya_atolye_adi", "TEXT"),
        ("degerlendirme_metodu", "TEXT DEFAULT 'FINE_KINNEY'"),
        ("frekans", "REAL"),
        ("risk_skoru", "REAL"),
        ("risk_seviyesi", "TEXT"),
        ("alinmasi_gereken_dof_onlem", "TEXT"),
        ("mebbis_kutuphane_kodu", "TEXT"),
        ("mebbis_dof_aciklama_metni", "TEXT"),
        ("odenek_talep_kapsaminda_mi", "INTEGER DEFAULT 0"),
        ("denetleyen_uzman", "TEXT DEFAULT 'Ramazan ERYILDIZ'"),
        ("ekip_onayi", "INTEGER DEFAULT 1")
    ]

    for col_name, col_type in new_risk_cols:
        if col_name not in existing_cols:
            c.execute(f"ALTER TABLE risk_assessments ADD COLUMN {col_name} {col_type}")
            print(f"[+] risk_assessments tablosuna '{col_name}' sütunu eklendi.")

    conn.commit()
    conn.close()
    print("[OK] Veritabanı şeması v2.6.0 başarıyla güncellendi.")

if __name__ == "__main__":
    upgrade_schema()

import sqlite3

def upgrade_schema():
    conn = sqlite3.connect('isg_muratpasa.db')
    c = conn.cursor()
    
    # Check risk_assessments columns
    c.execute("PRAGMA table_info(risk_assessments)")
    existing_cols = {row[1] for row in c.fetchall()}
    print(f"Mevcut sütun sayısı: {len(existing_cols)}")
    
    columns_to_add = [
        ("kaynak", "TEXT DEFAULT 'MEBBIS_RAPORU'"), # MEBBIS_RAPORU or SAHA_TESPITI
        ("saha_durumu", "TEXT DEFAULT 'TEYIT_BEKLIYOR'"), # TEYIT_BEKLIYOR, GIDERILDI, DEVAM_EDIYOR, KISMEN_GIDERILDI, YENI_TESPIT
        ("saha_notu", "TEXT"),
        ("saha_fotograf_yolu", "TEXT"),
        ("son_denetim_tarihi", "TEXT"),
        ("dof_no", "TEXT"),
        ("dof_acilis_tarihi", "TEXT"),
        ("dof_kapanis_tarihi", "TEXT"),
        ("mebbis_dosya_adi", "TEXT")
    ]
    
    for col_name, col_type in columns_to_add:
        if col_name not in existing_cols:
            print(f"Ekleniyor: {col_name} ({col_type})")
            c.execute(f"ALTER TABLE risk_assessments ADD COLUMN {col_name} {col_type}")
        else:
            print(f"Zaten mevcut: {col_name}")
            
    # Also create a table for tracking MEBBIS report file imports
    c.execute("""
        CREATE TABLE IF NOT EXISTS mebbis_imports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            school_id INTEGER NOT NULL,
            dosya_adi TEXT NOT NULL,
            toplam_madde INTEGER DEFAULT 0,
            yeni_eklenen INTEGER DEFAULT 0,
            guncellenen INTEGER DEFAULT 0,
            yukleyen_kullanici TEXT,
            yukleme_tarihi TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (school_id) REFERENCES schools(id)
        )
    """)
    print("mebbis_imports tablosu hazır.")
    
    conn.commit()
    conn.close()
    print("Veritabanı güncellemesi tamamlandı.")

if __name__ == '__main__':
    upgrade_schema()

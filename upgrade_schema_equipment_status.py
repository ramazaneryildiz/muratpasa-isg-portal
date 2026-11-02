import sqlite3

DB_PATH = r"c:\isg\isg_muratpasa.db"

def run_migration():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("PRAGMA table_info(inspection_metrics)")
    cols = [r[1] for r in c.fetchall()]

    new_cols = [
        ("kazan_durumu", "TEXT DEFAULT 'Var'"),
        ("asansor_durumu", "TEXT DEFAULT 'Yok'"),
        ("su_deposu_durumu", "TEXT DEFAULT 'Var'"),
        ("paratoner_durumu", "TEXT DEFAULT 'Var'")
    ]

    for col_name, col_type in new_cols:
        if col_name not in cols:
            c.execute(f"ALTER TABLE inspection_metrics ADD COLUMN {col_name} {col_type}")
            print(f"[+] inspection_metrics tablosuna '{col_name}' sütunu eklendi.")
        else:
            print(f"[-] '{col_name}' sütunu zaten mevcut.")

    conn.commit()
    conn.close()
    print("[OK] Veritabanı şema güncellemesi başarıyla tamamlandı.")

if __name__ == "__main__":
    run_migration()

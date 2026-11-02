import os
import sys
import sqlite3
import hashlib
import pypdf
import openpyxl
import re
import unicodedata

sys.stdout.reconfigure(encoding='utf-8')

DB_PATH = r"c:\isg\isg_muratpasa.db"

def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode('utf-8')).hexdigest()

def normalize(s):
    if not s: return ''
    s = str(s).lower().replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ş', 's').replace('ö', 'o').replace('ç', 'c')
    s = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in s if c.isalnum() or c.isspace()).strip()

def run_migration():
    print("=" * 65)
    print(" 1. TABLO SÜTUNLARI GENİŞLETİLİYOR")
    print("=" * 65)
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # 1. users tablosuna yeni sütunlar ekle
    c.execute("PRAGMA table_info(users)")
    existing_user_cols = [r[1] for r in c.fetchall()]
    user_cols = [
        ("tc_kimlik", "TEXT"),
        ("kadro_yeri", "TEXT"),
        ("igu_sinifi", "TEXT"),
        ("gorev_turu", "TEXT DEFAULT 'Tam Süreli'"),
        ("telefon", "TEXT")
    ]
    for col, ctype in user_cols:
        if col not in existing_user_cols:
            c.execute(f"ALTER TABLE users ADD COLUMN {col} {ctype}")
            print(f"[+] users tablosuna '{col}' sütunu eklendi.")

    # 2. schools tablosuna yeni sütunlar ekle
    c.execute("PRAGMA table_info(schools)")
    existing_school_cols = [r[1] for r in c.fetchall()]
    school_cols = [
        ("personel_sayisi", "INTEGER DEFAULT 0"),
        ("aylik_hizmet_suresi_dk", "INTEGER DEFAULT 0")
    ]
    for col, ctype in school_cols:
        if col not in existing_school_cols:
            c.execute(f"ALTER TABLE schools ADD COLUMN {col} {ctype}")
            print(f"[+] schools tablosuna '{col}' sütunu eklendi.")

    conn.commit()

    print("\n" + "=" * 65)
    print(" 2. 8 RESMÎ İŞ GÜVENLİĞİ UZMANI (İGU) SİSTEME EKLENİYOR")
    print("=" * 65)

    default_pw_hash = hash_pw("isg2026!")

    experts = [
        {
            "username": "ramazan.eryildiz",
            "full_name": "Ramazan ERYILDIZ",
            "role": "koordinator",
            "title": "İlçe İSG Büro Koordinatörü - A Sınıfı İGU",
            "email": "ramazaneryildiz@gmail.com",
            "tc_kimlik": "15157540378",
            "kadro_yeri": "Şehit Özcan Şenol MTAL (Kepez) / Muratpaşa İlçe MEM",
            "igu_sinifi": "A Sınıfı İGU",
            "gorev_turu": "Tam Süreli / Büro Yöneticisi & İGU"
        },
        {
            "username": "ahmet.aydogmus",
            "full_name": "Ahmet Tunca AYDOĞMUŞ",
            "role": "uzman",
            "title": "İş Güvenliği Uzmanı - C Sınıfı İGU",
            "email": "ahmet.aydogmus@meb.gov.tr",
            "tc_kimlik": "11758466772",
            "kadro_yeri": "Fettah Tamince Denizcilik MTAL (Konyaaltı)",
            "igu_sinifi": "C Sınıfı İGU",
            "gorev_turu": "Tam Süreli / İGU"
        },
        {
            "username": "birsen.dinc",
            "full_name": "Birsen ERSOY DİNÇ",
            "role": "uzman",
            "title": "İş Güvenliği Uzmanı - B Sınıfı İGU",
            "email": "birsen.dinc@meb.gov.tr",
            "tc_kimlik": "54310307250",
            "kadro_yeri": "Muratpaşa İlçe Millî Eğitim Müdürlüğü",
            "igu_sinifi": "B Sınıfı İGU",
            "gorev_turu": "Tam Süreli / İGU"
        },
        {
            "username": "bulent.acikgoz",
            "full_name": "Bülent AÇIKGÖZ",
            "role": "uzman",
            "title": "İş Güvenliği Uzmanı - C Sınıfı İGU",
            "email": "bulent.acikgoz@meb.gov.tr",
            "tc_kimlik": "5084881184",
            "kadro_yeri": "Akdeniz Sanayi Sitesi MTAL / Antalya İl MEM",
            "igu_sinifi": "C Sınıfı İGU",
            "gorev_turu": "Kısmi Süreli / İGU"
        },
        {
            "username": "havva.aksu",
            "full_name": "Havva AKSU",
            "role": "uzman",
            "title": "İş Güvenliği Uzmanı - A Sınıfı İGU",
            "email": "havva.aksu@meb.gov.tr",
            "tc_kimlik": "26818159730",
            "kadro_yeri": "Atatürk Ortaokulu / Muratpaşa İlçe MEM",
            "igu_sinifi": "A Sınıfı İGU",
            "gorev_turu": "Tam Süreli / İGU"
        },
        {
            "username": "mete.ozdemir",
            "full_name": "Mete ÖZDEMİR",
            "role": "uzman",
            "title": "İş Güvenliği Uzmanı - C Sınıfı İGU",
            "email": "mete.ozdemir@meb.gov.tr",
            "tc_kimlik": "45097113516",
            "kadro_yeri": "Antalya İl Millî Eğitim Müdürlüğü",
            "igu_sinifi": "C Sınıfı İGU",
            "gorev_turu": "Tam Süreli / İGU"
        },
        {
            "username": "nazim.ozbakir",
            "full_name": "Nazım ÖZBAKIR",
            "role": "uzman",
            "title": "İş Güvenliği Uzmanı - A Sınıfı İGU",
            "email": "nazim.ozbakir@meb.gov.tr",
            "tc_kimlik": "52570328350",
            "kadro_yeri": "Düden MTAL (Kepez)",
            "igu_sinifi": "A Sınıfı İGU",
            "gorev_turu": "Kısmi Süreli / İGU"
        },
        {
            "username": "serdar.basmaci",
            "full_name": "Serdar BASMACI",
            "role": "uzman",
            "title": "İş Güvenliği Uzmanı - C Sınıfı İGU",
            "email": "serdar.basmaci@meb.gov.tr",
            "tc_kimlik": "32230380416",
            "kadro_yeri": "Antalya İl Millî Eğitim Müdürlüğü",
            "igu_sinifi": "C Sınıfı İGU",
            "gorev_turu": "Tam Süreli / İGU"
        }
    ]

    # Delete placeholder users like uzman.ahmet and uzman.fatma
    c.execute("DELETE FROM users WHERE username IN ('uzman.ahmet', 'uzman.fatma')")

    expert_id_map = {}
    for exp in experts:
        existing = c.execute("SELECT id FROM users WHERE username = ?", (exp["username"],)).fetchone()
        if existing:
            c.execute("""
                UPDATE users
                SET full_name = ?, role = ?, title = ?, email = ?,
                    tc_kimlik = ?, kadro_yeri = ?, igu_sinifi = ?, gorev_turu = ?
                WHERE id = ?
            """, (
                exp["full_name"], exp["role"], exp["title"], exp["email"],
                exp["tc_kimlik"], exp["kadro_yeri"], exp["igu_sinifi"], exp["gorev_turu"],
                existing[0]
            ))
            expert_id_map[exp["full_name"]] = existing[0]
            print(f"[OK] Güncellendi: {exp['full_name']} (ID: {existing[0]})")
        else:
            c.execute("""
                INSERT INTO users (username, password_hash, full_name, role, title, email, tc_kimlik, kadro_yeri, igu_sinifi, gorev_turu)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                exp["username"], default_pw_hash, exp["full_name"], exp["role"], exp["title"], exp["email"],
                exp["tc_kimlik"], exp["kadro_yeri"], exp["igu_sinifi"], exp["gorev_turu"]
            ))
            new_id = c.lastrowid
            expert_id_map[exp["full_name"]] = new_id
            print(f"[OK] Yeni Eklendi: {exp['full_name']} (ID: {new_id})")

    conn.commit()

    print("\n" + "=" * 65)
    print(" 3. İLÇE VE İL GENELİ GÖREVLENDİRMELERİ VERİTABANINA İŞLENİYOR")
    print("=" * 65)

    # 1. Read PDF for employee count and service duration
    pdf_path = r"c:\isg\IL GENEL IGU OKUL VE PERSONEL GOREVLENDIRME 20.08.2026.pdf"
    reader = pypdf.PdfReader(pdf_path)
    pdf_data = {}
    for page in reader.pages:
        for line in page.extract_text().split('\n'):
            if 'MURATPAŞA' in line.upper() or 'MURATPASA' in line.upper():
                m = re.search(r'(\d+)\s+MURATPA[SŞ]A\s+(\d{6})\s+(.+?)\s+(Az Tehlikeli|Tehlikeli|Çok Tehlikeli)\s+(\d+)\s+(\d+)', line)
                if m:
                    sno = int(m.group(1))
                    kkod = m.group(2)
                    okul_adi = m.group(3).strip()
                    tehlike = m.group(4)
                    personel = int(m.group(5))
                    sure_dk = int(m.group(6))
                    pdf_data[kkod] = {
                        "sno": sno,
                        "tehlike": tehlike,
                        "personel": personel,
                        "sure_dk": sure_dk
                    }

    print(f"[+] PDF'den {len(pdf_data)} Muratpaşa kurumunun personel ve süre bilgisi okundu.")

    # 2. Map SNO ranges from PDF & Excel to experts
    # SNO ranges:
    # 173 - 179: Mete ÖZDEMİR
    # 262 - 269: Birsen ERSOY DİNÇ
    # 270 - 277: Bülent AÇIKGÖZ
    # 278 - 282: Havva AKSU
    # 283 - 289: Ramazan ERYILDIZ
    # 292 - 294: Nazım ÖZBAKIR
    # 295 - 303: Serdar BASMACI
    # 304 - 317: Ahmet Tunca AYDOĞMUŞ

    def get_expert_for_sno(sno):
        if 173 <= sno <= 179:
            return "Mete ÖZDEMİR"
        elif 262 <= sno <= 269:
            return "Birsen ERSOY DİNÇ"
        elif 270 <= sno <= 277:
            return "Bülent AÇIKGÖZ"
        elif 278 <= sno <= 282:
            return "Havva AKSU"
        elif 283 <= sno <= 289:
            return "Ramazan ERYILDIZ"
        elif 292 <= sno <= 294:
            return "Nazım ÖZBAKIR"
        elif 295 <= sno <= 303:
            return "Serdar BASMACI"
        elif 304 <= sno <= 317:
            return "Ahmet Tunca AYDOĞMUŞ"
        return None

    # Apply PDF assignments by kurum_kodu
    assigned_count = 0
    for kkod, info in pdf_data.items():
        exp_name = get_expert_for_sno(info["sno"])
        if exp_name and exp_name in expert_id_map:
            u_id = expert_id_map[exp_name]
            c.execute("""
                UPDATE schools
                SET sorumlu_uzman_id = ?,
                    personel_sayisi = ?,
                    aylik_hizmet_suresi_dk = ?,
                    tehlike_sinifi = ?
                WHERE kurum_kodu = ?
            """, (u_id, info["personel"], info["sure_dk"], info["tehlike"], kkod))
            if c.rowcount > 0:
                assigned_count += 1

    conn.commit()
    print(f"[+] Toplam {assigned_count} resmî okul kurum kodları üzerinden ilgili İGU'ya atandı.")

    # 3. Check for any remaining schools from the Excel file (e.g. combined campuses)
    excel_files = [f for f in os.listdir(r"c:\isg") if 'GOREV' in f.upper() and f.endswith('.xlsx')]
    wb = openpyxl.load_workbook(os.path.join(r"c:\isg", excel_files[0]), data_only=True)
    sheet = wb.active

    schools_db = c.execute("SELECT id, kurum_kodu, kurum_adi, sorumlu_uzman_id FROM schools").fetchall()
    db_school_map = {normalize(s[2]): s for s in schools_db}

    current_expert_name = None
    excel_matched = 0

    for r in range(4, sheet.max_row + 1):
        name_cell = sheet.cell(r, 3).value
        school_cell = sheet.cell(r, 6).value

        if name_cell:
            current_expert_name = str(name_cell).strip()

        if school_cell and current_expert_name:
            target_school_name = str(school_cell).strip()
            norm_name = normalize(target_school_name)
            
            # Find in DB
            target_id = None
            if norm_name in db_school_map:
                target_id = db_school_map[norm_name][0]
            else:
                for db_norm, db_s in db_school_map.items():
                    tokens = [t for t in norm_name.split() if len(t) > 3 and t not in ['okulu', 'ortaokulu', 'ilkokulu', 'lisesi', 'anadolu', 'mesleki', 'teknik']]
                    if tokens and all(t in db_norm for t in tokens[:2]):
                        target_id = db_s[0]
                        break

            if target_id and current_expert_name in expert_id_map:
                u_id = expert_id_map[current_expert_name]
                c.execute("UPDATE schools SET sorumlu_uzman_id = ? WHERE id = ? AND (sorumlu_uzman_id IS NULL OR sorumlu_uzman_id != ?)", (u_id, target_id, u_id))
                if c.rowcount > 0:
                    excel_matched += 1

    conn.commit()

    print("\n" + "=" * 65)
    print(" 4. ATAMA SONUÇLARI VE ÖZET TABLO")
    print("=" * 65)

    c.execute("""
        SELECT u.id, u.full_name, u.igu_sinifi, u.kadro_yeri,
               COUNT(s.id) as okul_sayisi,
               SUM(s.personel_sayisi) as toplam_personel,
               SUM(s.aylik_hizmet_suresi_dk) as toplam_dakika
        FROM users u
        LEFT JOIN schools s ON u.id = s.sorumlu_uzman_id
        GROUP BY u.id
        ORDER BY okul_sayisi DESC
    """)

    for row in c.fetchall():
        uid, fname, sinif, kadro, okul_cnt, pers_cnt, dk_cnt = row
        dk_str = f"{dk_cnt} dk ({dk_cnt//60} sa {dk_cnt%60} dk)" if dk_cnt else "0 dk"
        print(f"[{uid}] {fname:22} ({sinif}) -> {okul_cnt:2} Okul | {pers_cnt or 0:4} Personel | Hizmet Süresi: {dk_str}")

    conn.close()

if __name__ == "__main__":
    run_migration()

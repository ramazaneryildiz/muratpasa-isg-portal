import os
import sys
import sqlite3
import unicodedata
import openpyxl
import xlrd

sys.stdout.reconfigure(encoding='utf-8')

DB_PATH = r"c:\isg\isg_muratpasa.db"
ISG_DIR = r"C:\iş_güvenliği"

def clean_phone(val):
    if not val:
        return ""
    s = str(val).split('.')[0].strip()
    digits = ''.join(c for c in s if c.isdigit())
    if len(digits) == 10 and digits.startswith('5'):
        return '0' + digits
    elif len(digits) == 11 and digits.startswith('05'):
        return digits
    elif digits:
        return digits
    return ""

def normalize(s):
    if not s:
        return ""
    s = str(s).lower().replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ş', 's').replace('ö', 'o').replace('ç', 'c')
    s = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in s if c.isalnum() or c.isspace()).strip()

def import_school_contacts(conn):
    print("=" * 65)
    print(" 1. OKUL MÜDÜRÜ VE İSG MÜDÜR YARDIMCISI İLETİŞİM BİLGİLERİ AKTARILIYOR")
    print("=" * 65)
    
    # Check for the file in c:\isg
    xlsx_files = [f for f in os.listdir(r"c:\isg") if f.endswith(".xlsx") and "veren" in f]
    if not xlsx_files:
        print("[-] İşveren bilgileri Excel dosyası bulunamadı.")
        return

    excel_path = os.path.join(r"c:\isg", xlsx_files[0])
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = wb.active
    
    cursor = conn.cursor()
    schools = cursor.execute("SELECT id, kurum_kodu, kurum_adi FROM schools").fetchall()
    school_map = {normalize(s[2]): s for s in schools}
    
    updated = 0
    for r in range(8, sheet.max_row + 1):
        name = sheet.cell(r, 4).value
        p_name = sheet.cell(r, 5).value
        p_phone = clean_phone(sheet.cell(r, 7).value)
        vp_name = sheet.cell(r, 8).value
        vp_phone = clean_phone(sheet.cell(r, 10).value)
        
        if not name or not p_name:
            continue
            
        norm_name = normalize(name)
        target_id = None
        
        if norm_name in school_map:
            target_id = school_map[norm_name][0]
        else:
            for db_norm, db_s in school_map.items():
                tokens = [t for t in norm_name.split() if len(t) > 3 and t not in ['okulu', 'mudurlugu', 'ortaokulu', 'ilkokulu', 'lisesi', 'anadolu', 'mesleki', 'teknik']]
                if tokens and all(t in db_norm for t in tokens[:2]):
                    target_id = db_s[0]
                    break
                    
        if target_id:
            cursor.execute("""
                UPDATE schools 
                SET mudur_ad_soyad = ?, 
                    mudur_tel = ?, 
                    isveren_vekili_ad_soyad = ?, 
                    isveren_vekili_tel = ?
                WHERE id = ?
            """, (
                str(p_name).strip(), 
                p_phone, 
                str(vp_name).strip() if vp_name else "", 
                vp_phone, 
                target_id
            ))
            updated += 1
            
    conn.commit()
    print(f"[+] Toplam {updated} okulun Müdür ve İSG Müdür Yrd. bilgileri ve telefonları güncellendi.")

def update_priority_degrees(conn):
    print("\n" + "=" * 65)
    print(" 2. RESMÎ ÖNCELİK DERECELERİ GÜNCELLENİYOR (102 OKUL)")
    print("=" * 65)
    fp = os.path.join(ISG_DIR, "ÖNCELİKLİ OKULLAR LİSTESİ ANTALYA İL GENELİ2026.xlsx")
    if not os.path.exists(fp):
        print("[-] Öncelikli okullar dosyası bulunamadı.")
        return

    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb['İL GENELİ']
    cursor = conn.cursor()

    updated = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[2]:
            continue
        ilce = str(row[2]).strip().upper()
        if "MURATPAŞA" in ilce or "MURATPASA" in ilce:
            kkod = str(row[3]).strip()
            derece = row[10]
            try:
                derece_int = int(derece)
            except Exception:
                derece_int = 3

            cursor.execute("""
                UPDATE schools SET oncelik_derecesi = ? WHERE kurum_kodu = ?
            """, (derece_int, kkod))
            if cursor.rowcount > 0:
                updated += 1

    conn.commit()
    print(f"[+] Toplam {updated} Muratpaşa okulunun resmî MEB/Emniyet öncelik derecesi güncellendi.")

def import_all_risk_reports(conn):
    print("\n" + "=" * 65)
    print(" 3. 6 BÜYÜK MTAL RİSK RAPORLARI VERİTABANINA AKTARILIYOR (MEBBİS)")
    print("=" * 65)
    rr_dir = os.path.join(ISG_DIR, "risk_raporları")
    if not os.path.exists(rr_dir):
        print("[-] Risk raporları klasörü bulunamadı.")
        return

    cursor = conn.cursor()
    cursor.execute("DELETE FROM risk_assessments")

    # Map file substring / institution code to school id
    file_school_map = {
        "102741": "970149", # Antalya MTAL
        "103328": "969922", # Antalya Ticaret Borsası Ticaret MTAL
        "103433": "970151", # Tez-Tur Turizm MTAL
        "103529": "751351", # Necati Dölen Ticaret MTAL
        "103629": "764955", # Muratpaşa MTAL
        "103713": "969926"  # Şehit Serkan Çölkesen Ticaret MTAL
    }

    cursor.execute("SELECT id, kurum_kodu, kurum_adi FROM schools")
    code_to_school = {row[1]: (row[0], row[2]) for row in cursor.fetchall()}

    total_inserted = 0

    for f in sorted(os.listdir(rr_dir)):
        if not f.endswith(".xls"):
            continue
            
        matched_code = None
        for key, scode in file_school_map.items():
            if key in f:
                matched_code = scode
                break
                
        if not matched_code or matched_code not in code_to_school:
            print(f"[-] Eşleşen okul kodu bulunamadı: {f}")
            continue

        school_id, school_name = code_to_school[matched_code]
        fp = os.path.join(rr_dir, f)
        
        try:
            book = xlrd.open_workbook(fp)
            sheet = book.sheet_by_index(0)

            file_inserted = 0
            for r in range(sheet.nrows):
                c0 = str(sheet.cell_value(r, 0)).strip() if sheet.ncols > 0 else ""
                if "BRK" not in c0:
                    continue
                    
                barkod = c0
                rev = str(sheet.cell_value(r, 2)).strip() if sheet.ncols > 2 else "1.0"
                birim = str(sheet.cell_value(r, 3)).strip() if sheet.ncols > 3 else ""
                tehlike = str(sheet.cell_value(r, 4)).strip() if sheet.ncols > 4 else ""
                tehlike_unsuru = str(sheet.cell_value(r, 6)).strip() if sheet.ncols > 6 else ""
                risk = str(sheet.cell_value(r, 9)).strip() if sheet.ncols > 9 else ""
                olasilik = str(sheet.cell_value(r, 11)).strip() if sheet.ncols > 11 else "2. Düşük"
                siddet = str(sheet.cell_value(r, 13)).strip() if sheet.ncols > 13 else "3. Orta"
                
                puan_raw = sheet.cell_value(r, 14) if sheet.ncols > 14 else 6
                try:
                    puan = int(float(puan_raw))
                except Exception:
                    puan = 6
                    
                sorumlu = str(sheet.cell_value(r, 15)).strip() if sheet.ncols > 15 else "İşveren ve İşveren Vekili"
                onlem = str(sheet.cell_value(r, 17)).strip() if sheet.ncols > 17 else ""
                mevzuat = str(sheet.cell_value(r, 20)).strip() if sheet.ncols > 20 else ""
                termin = str(sheet.cell_value(r, 23)).strip() if sheet.ncols > 23 else ""
                
                # Belirleme DÖF durumu
                dof = "Devam Ediyor"
                if "tamamlandı" in onlem.lower() or "yapıldı" in onlem.lower():
                    dof = "Kapatıldı"

                cursor.execute("""
                INSERT INTO risk_assessments (
                    school_id, barkod_no, revizyon_no, kontrol_edilen_birim, tehlike,
                    tehlike_unsurlari, riskler, olasilik, siddet, risk_puani,
                    sorumlular, alinan_onlemler, dayanak_mevzuat, termin_tarihi, dof_durumu
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                """, (
                    school_id, barkod, rev, birim, tehlike,
                    tehlike_unsuru, risk, olasilik, siddet, puan,
                    sorumlu, onlem, mevzuat, termin, dof
                ))
                file_inserted += 1

            total_inserted += file_inserted
            print(f"[+] {school_name[:35]:35} -> {file_inserted} MEBBİS risk kaydı yüklendi.")

        except Exception as e:
            print(f"[-] {f} okunurken hata: {e}")

    conn.commit()
    print(f"\n[SUCCESS] Toplam aktarılan gerçek MEBBİS risk analizi kaydı: {total_inserted}")

def main():
    conn = sqlite3.connect(DB_PATH)
    import_school_contacts(conn)
    update_priority_degrees(conn)
    import_all_risk_reports(conn)
    
    # Final check
    cursor = conn.cursor()
    cnt_phones = cursor.execute("SELECT COUNT(*) FROM schools WHERE mudur_tel IS NOT NULL AND mudur_tel != ''").fetchone()[0]
    cnt_prio = cursor.execute("SELECT COUNT(*) FROM schools WHERE oncelik_derecesi != 3").fetchone()[0]
    cnt_risks = cursor.execute("SELECT COUNT(*) FROM risk_assessments").fetchone()[0]
    
    print("\n" + "=" * 65)
    print(" VERİTABANI GÜNCELLEME ÖZETİ")
    print("=" * 65)
    print(f" İletişim Bilgisi Eklenen Okul Sayısı : {cnt_phones} / 283")
    print(f" Resmî Öncelik Derecesi Olan Okullar  : {cnt_prio}")
    print(f" Veritabanındaki Toplam Risk Kaydı    : {cnt_risks}")
    print("=" * 65)
    
    conn.close()

if __name__ == "__main__":
    main()

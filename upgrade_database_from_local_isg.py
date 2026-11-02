import os
import sys
import sqlite3
import openpyxl
import xlrd
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

DB_PATH = r"c:\isg\isg_muratpasa.db"
ISG_DIR = r"C:\iş_güvenliği"

def update_priority_degrees(conn):
    print("=" * 65)
    print(" 1. RESMÎ ÖNCELİK DERECELERİ GÜNCELLENİYOR (102 OKUL)")
    print("=" * 65)
    fp = os.path.join(ISG_DIR, "ÖNCELİKLİ OKULLAR LİSTESİ ANTALYA İL GENELİ2026.xlsx")
    if not os.path.exists(fp):
        print("Öncelikli okullar dosyası bulunamadı.")
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
    print(f"[+] Toplam {updated} Muratpaşa okulunun resmî öncelik derecesi güncellendi.")

def import_all_risk_reports(conn):
    print("\n" + "=" * 65)
    print(" 2. 6 BÜYÜK MTAL RİSK RAPORLARI VERİTABANINA AKTARILIYOR (3.100+ SATIR)")
    print("=" * 65)
    rr_dir = os.path.join(ISG_DIR, "risk_raporları")
    if not os.path.exists(rr_dir):
        return

    cursor = conn.cursor()

    # Önce mevcut risk_assessments tablosunu temizleyelim ki çift kayıt olmasın
    cursor.execute("DELETE FROM risk_assessments")

    school_name_to_code = {
        "Antalya Mesleki ve Teknik Anadolu Lisesi": "970149",
        "Antalya Ticaret Borsası Ticaret Mesleki ve Teknik Anadolu Lisesi": "969922",
        "Tez-Tur Turizm Mesleki ve Teknik Anadolu Lisesi": "970151",
        "Necati Dölen Ticaret Mesleki ve Teknik Anadolu Lisesi": "751351",
        "Muratpaşa Mesleki ve Teknik Anadolu Lisesi": "764955",
        "Şehit Serkan Çölkesen Ticaret Mesleki ve Teknik Anadolu Lisesi": "969926"
    }

    # Okul kodundan school_id alalım
    cursor.execute("SELECT id, kurum_kodu FROM schools")
    code_to_id = {row[1]: row[0] for row in cursor.fetchall()}

    total_inserted = 0

    for f in os.listdir(rr_dir):
        if not f.endswith(".xls"):
            continue
        fp = os.path.join(rr_dir, f)
        try:
            book = xlrd.open_workbook(fp)
            sheet = book.sheet_by_index(0)

            # Okul adını bulalım
            matched_school_id = None
            for r in range(min(10, sheet.nrows)):
                for c in range(min(5, sheet.ncols)):
                    val = str(sheet.cell_value(r, c)).strip()
                    for sname, scode in school_name_to_code.items():
                        if sname.lower() in val.lower() or val.lower() in sname.lower():
                            matched_school_id = code_to_id.get(scode)
                            break
                    if matched_school_id:
                        break
                if matched_school_id:
                    break

            if not matched_school_id:
                print(f"[-] Eşleşen okul bulunamadı: {f}")
                continue

            # Başlık satırını bul (genellikle 5. veya 6. satırda 'Barkod No' veya 'Birim' olur)
            header_row = None
            for r in range(min(15, sheet.nrows)):
                row_vals = [str(sheet.cell_value(r, c)).strip().lower() for c in range(sheet.ncols)]
                if any("barkod" in v for v in row_vals) or any("tehlike" in v for v in row_vals):
                    header_row = r
                    break

            if header_row is None:
                header_row = 5

            file_inserted = 0
            for r in range(header_row + 1, sheet.nrows):
                if sheet.ncols < 10:
                    continue
                barkod = str(sheet.cell_value(r, 0)).strip() if sheet.ncols > 0 else ""
                rev = str(sheet.cell_value(r, 1)).strip() if sheet.ncols > 1 else "1"
                birim = str(sheet.cell_value(r, 2)).strip() if sheet.ncols > 2 else ""
                tehlike = str(sheet.cell_value(r, 3)).strip() if sheet.ncols > 3 else ""
                tehlike_unsuru = str(sheet.cell_value(r, 4)).strip() if sheet.ncols > 4 else ""
                risk = str(sheet.cell_value(r, 5)).strip() if sheet.ncols > 5 else ""
                olasilik = str(sheet.cell_value(r, 6)).strip() if sheet.ncols > 6 else "2. Düşük"
                siddet = str(sheet.cell_value(r, 7)).strip() if sheet.ncols > 7 else "3. Orta"
                puan_raw = sheet.cell_value(r, 8) if sheet.ncols > 8 else 6
                try:
                    puan = int(float(puan_raw))
                except Exception:
                    puan = 6
                sorumlu = str(sheet.cell_value(r, 9)).strip() if sheet.ncols > 9 else "İŞVEREN VEKİLİ"
                onlem = str(sheet.cell_value(r, 10)).strip() if sheet.ncols > 10 else ""
                mevzuat = str(sheet.cell_value(r, 11)).strip() if sheet.ncols > 11 else ""
                termin = str(sheet.cell_value(r, 12)).strip() if sheet.ncols > 12 else ""
                dof = str(sheet.cell_value(r, 13)).strip() if sheet.ncols > 13 else "Süresi Devam Ediyor"

                if not tehlike and not birim:
                    continue

                cursor.execute("""
                INSERT INTO risk_assessments (
                    school_id, barkod_no, revizyon_no, kontrol_edilen_birim, tehlike,
                    tehlike_unsurlari, riskler, olasilik, siddet, risk_puani,
                    sorumlular, alinan_onlemler, dayanak_mevzuat, termin_tarihi, dof_durumu
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                """, (
                    matched_school_id, barkod, rev, birim, tehlike,
                    tehlike_unsuru, risk, str(olasilik), str(siddet), puan,
                    sorumlu, onlem, mevzuat, termin, dof
                ))
                file_inserted += 1

            total_inserted += file_inserted
            print(f"[+] {f} -> Okul ID {matched_school_id}: {file_inserted} risk kaydı aktarıldı.")

        except Exception as e:
            print(f"[-] {f} işlenirken hata: {e}")

    conn.commit()
    print(f"\n[SUCCESS] Toplam aktarılan gerçek MEBBİS risk kaydı: {total_inserted}")

def main():
    conn = sqlite3.connect(DB_PATH)
    update_priority_degrees(conn)
    import_all_risk_reports(conn)
    conn.close()

if __name__ == "__main__":
    main()

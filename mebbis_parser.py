"""
MEBBİS İSG Risk Değerlendirme Raporu Ayrıştırma ve İçe Aktarma Motoru
Muratpaşa İlçe MEM İSG Portalı - 2026
"""

import os
import re
import sqlite3
import datetime
from typing import Dict, List, Any, Optional

try:
    import xlrd
except ImportError:
    xlrd = None

try:
    import openpyxl
except ImportError:
    openpyxl = None


def normalize_text(text: str) -> str:
    """Türkçe karakterleri ve boşlukları normalize eder."""
    if not text:
        return ""
    t = text.strip()
    return re.sub(r'\s+', ' ', t)


def calculate_risk_level(score: float) -> str:
    """5x5 MEBBİS Matris veya Fine-Kinney puanına göre seviye belirler."""
    if score >= 20:
        return "ÇOK YÜKSEK"
    elif score >= 15:
        return "YÜKSEK"
    elif score >= 8:
        return "ORTA"
    else:
        return "DÜŞÜK"


def parse_mebbis_xls(file_path: str) -> Dict[str, Any]:
    """MEBBİS .xls formatındaki raporu okur ve risk maddelerini listeler."""
    if not xlrd:
        raise RuntimeError("xlrd kütüphanesi kurulu değil.")

    book = xlrd.open_workbook(file_path)
    sheet = book.sheet_by_index(0)

    # 1. Okul ve İlçe Bilgisini Bul
    ilce = "MURATPAŞA"
    kurum_adi = ""
    for r in range(min(6, sheet.nrows)):
        for c in range(sheet.ncols):
            val = str(sheet.cell_value(r, c)).strip()
            if "KURUM ADI" in val.upper():
                # bir sonraki veya 2. sütundaki değer
                for next_c in range(c + 1, min(c + 4, sheet.ncols)):
                    nv = str(sheet.cell_value(r, next_c)).strip()
                    if nv and nv != ":":
                        kurum_adi = nv
                        break
            if "İLÇESİ" in val.upper():
                for next_c in range(c + 1, min(c + 4, sheet.ncols)):
                    nv = str(sheet.cell_value(r, next_c)).strip()
                    if nv and nv != ":":
                        ilce = nv
                        break

    # 2. Başlık Satırını Bul
    header_row = -1
    for r in range(min(15, sheet.nrows)):
        for c in range(sheet.ncols):
            if "Barkod" in str(sheet.cell_value(r, c)):
                header_row = r
                break
        if header_row != -1:
            break

    if header_row == -1:
        header_row = 6  # MEBBİS varsayılanı

    items: List[Dict[str, Any]] = []
    extracted_kurum_kodlari = set()

    for r in range(header_row + 1, sheet.nrows):
        col0 = str(sheet.cell_value(r, 0)).strip()
        if not col0 or "BRK" not in col0.upper():
            continue

        barcode = col0
        # Kurum kodunu barkodun ilk 6 hanesinden çıkar
        m = re.match(r"^(\d{6})", barcode)
        if m:
            extracted_kurum_kodlari.add(m.group(1))

        # Revizyon
        rev_raw = str(sheet.cell_value(r, 2)).strip()
        try:
            rev_no = str(int(float(rev_raw)))
        except (ValueError, TypeError):
            rev_no = rev_raw

        unit = normalize_text(str(sheet.cell_value(r, 3)))
        hazard = normalize_text(str(sheet.cell_value(r, 4)))
        hazard_elements = normalize_text(str(sheet.cell_value(r, 6))) if sheet.ncols > 6 else ""
        risks = normalize_text(str(sheet.cell_value(r, 9))) if sheet.ncols > 9 else ""
        prob = normalize_text(str(sheet.cell_value(r, 11))) if sheet.ncols > 11 else ""
        sev = normalize_text(str(sheet.cell_value(r, 13))) if sheet.ncols > 13 else ""

        # Puan
        score_raw = str(sheet.cell_value(r, 14)).strip() if sheet.ncols > 14 else "0"
        try:
            score = float(score_raw)
        except (ValueError, TypeError):
            score = 0.0

        responsibles = normalize_text(str(sheet.cell_value(r, 15))) if sheet.ncols > 15 else ""
        measures = normalize_text(str(sheet.cell_value(r, 17))) if sheet.ncols > 17 else ""
        legislation = normalize_text(str(sheet.cell_value(r, 20))) if sheet.ncols > 20 else ""

        # Termin Süresi
        deadline = ""
        if sheet.ncols > 23:
            raw_deadline = sheet.cell_value(r, 23)
            if sheet.cell_type(r, 23) == xlrd.XL_CELL_DATE:
                try:
                    dt_tuple = xlrd.xldate_as_tuple(raw_deadline, book.datemode)
                    deadline = f"{dt_tuple[2]:02d}/{dt_tuple[1]:02d}/{dt_tuple[0]}"
                except Exception:
                    deadline = str(raw_deadline).strip()
            else:
                deadline = str(raw_deadline).strip()

        level = calculate_risk_level(score)

        items.append({
            "barkod_no": barcode,
            "revizyon_no": rev_no,
            "kontrol_edilen_birim": unit,
            "bolum_unite": unit,
            "tehlike": hazard,
            "tehlike_unsurlari": hazard_elements,
            "riskler": risks,
            "olasilik": prob,
            "siddet": sev,
            "risk_puani": int(score) if score.is_integer() else score,
            "risk_skoru": score,
            "risk_seviyesi": level,
            "sorumlular": responsibles,
            "alinan_onlemler": measures,
            "alinmasi_gereken_dof_onlem": measures,
            "dayanak_mevzuat": legislation,
            "termin_tarihi": deadline,
            "kaynak": "MEBBIS_RAPORU",
            "saha_durumu": "TEYIT_BEKLIYOR",
            "dof_durumu": "Devam Ediyor"
        })

    primary_kurum_kodu = list(extracted_kurum_kodlari)[0] if extracted_kurum_kodlari else ""

    return {
        "dosya_adi": os.path.basename(file_path),
        "kurum_adi": kurum_adi,
        "ilce": ilce,
        "kurum_kodu": primary_kurum_kodu,
        "toplam_madde": len(items),
        "items": items
    }


def parse_mebbis_xlsx(file_path: str) -> Dict[str, Any]:
    """MEBBİS .xlsx formatındaki raporu openpyxl ile okur."""
    if not openpyxl:
        raise RuntimeError("openpyxl kütüphanesi kurulu değil.")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active

    kurum_adi = ""
    ilce = "MURATPAŞA"

    for r in range(1, min(7, sheet.max_row + 1)):
        for c in range(1, min(10, sheet.max_column + 1)):
            val = str(sheet.cell(r, c).value or "").strip()
            if "KURUM ADI" in val.upper():
                for next_c in range(c + 1, min(c + 4, sheet.max_column + 1)):
                    nv = str(sheet.cell(r, next_c).value or "").strip()
                    if nv and nv != ":":
                        kurum_adi = nv
                        break

    items: List[Dict[str, Any]] = []
    extracted_kurum_kodlari = set()

    for r in range(7, sheet.max_row + 1):
        col0 = str(sheet.cell(r, 1).value or "").strip()
        if not col0 or "BRK" not in col0.upper():
            continue

        barcode = col0
        m = re.match(r"^(\d{6})", barcode)
        if m:
            extracted_kurum_kodlari.add(m.group(1))

        rev_raw = str(sheet.cell(r, 3).value or "").strip()
        try:
            rev_no = str(int(float(rev_raw)))
        except (ValueError, TypeError):
            rev_no = rev_raw

        unit = normalize_text(str(sheet.cell(r, 4).value or ""))
        hazard = normalize_text(str(sheet.cell(r, 5).value or ""))
        hazard_elements = normalize_text(str(sheet.cell(r, 7).value or ""))
        risks = normalize_text(str(sheet.cell(r, 10).value or ""))
        prob = normalize_text(str(sheet.cell(r, 12).value or ""))
        sev = normalize_text(str(sheet.cell(r, 14).value or ""))

        score_raw = str(sheet.cell(r, 15).value or "0").strip()
        try:
            score = float(score_raw)
        except (ValueError, TypeError):
            score = 0.0

        responsibles = normalize_text(str(sheet.cell(r, 16).value or ""))
        measures = normalize_text(str(sheet.cell(r, 18).value or ""))
        legislation = normalize_text(str(sheet.cell(r, 21).value or ""))
        
        raw_deadline = sheet.cell(r, 24).value
        if isinstance(raw_deadline, (datetime.datetime, datetime.date)):
            deadline = raw_deadline.strftime("%d/%m/%Y")
        else:
            deadline = str(raw_deadline or "").strip()

        level = calculate_risk_level(score)

        items.append({
            "barkod_no": barcode,
            "revizyon_no": rev_no,
            "kontrol_edilen_birim": unit,
            "bolum_unite": unit,
            "tehlike": hazard,
            "tehlike_unsurlari": hazard_elements,
            "riskler": risks,
            "olasilik": prob,
            "siddet": sev,
            "risk_puani": int(score) if score.is_integer() else score,
            "risk_skoru": score,
            "risk_seviyesi": level,
            "sorumlular": responsibles,
            "alinan_onlemler": measures,
            "alinmasi_gereken_dof_onlem": measures,
            "dayanak_mevzuat": legislation,
            "termin_tarihi": deadline,
            "kaynak": "MEBBIS_RAPORU",
            "saha_durumu": "TEYIT_BEKLIYOR",
            "dof_durumu": "Devam Ediyor"
        })

    primary_kurum_kodu = list(extracted_kurum_kodlari)[0] if extracted_kurum_kodlari else ""

    return {
        "dosya_adi": os.path.basename(file_path),
        "kurum_adi": kurum_adi,
        "ilce": ilce,
        "kurum_kodu": primary_kurum_kodu,
        "toplam_madde": len(items),
        "items": items
    }


def match_school(conn: sqlite3.Connection, kurum_kodu: str, kurum_adi: str) -> Optional[sqlite3.Row]:
    """Verilen kod veya ada göre schools tablosundan okulu bulur."""
    c = conn.cursor()
    # 1. Kodla tam eşleşme
    if kurum_kodu:
        c.execute("SELECT * FROM schools WHERE kurum_kodu = ?", (kurum_kodu,))
        s = c.fetchone()
        if s:
            return s

    # 2. İsimle eşleşme
    if kurum_adi:
        c.execute("SELECT * FROM schools WHERE UPPER(kurum_adi) = UPPER(?)", (kurum_adi.strip(),))
        s = c.fetchone()
        if s:
            return s

        # Like eşleştirme
        clean_name = kurum_adi.replace("Mesleki ve Teknik Anadolu Lisesi", "").replace("MTAL", "").strip()
        if clean_name:
            c.execute("SELECT * FROM schools WHERE kurum_adi LIKE ?", (f"%{clean_name}%",))
            matches = c.fetchall()
            if len(matches) == 1:
                return matches[0]

    return None


def import_mebbis_parsed_data(conn: sqlite3.Connection, parsed: Dict[str, Any], explicit_school_id: Optional[int] = None, user_name: str = "Ramazan ERYILDIZ") -> Dict[str, Any]:
    """Ayrıştırılan veriyi isg_muratpasa.db veritabanına aktarır."""
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    school = None
    if explicit_school_id:
        c.execute("SELECT * FROM schools WHERE id = ?", (explicit_school_id,))
        school = c.fetchone()

    if not school:
        school = match_school(conn, parsed.get("kurum_kodu", ""), parsed.get("kurum_adi", ""))

    if not school:
        return {
            "success": False,
            "error": f"Okul eşleştirilemedi. Dosyadaki Kurum Kodu: {parsed.get('kurum_kodu')}, Kurum Adı: {parsed.get('kurum_adi')}",
            "parsed": parsed
        }

    school_id = school["id"]
    school_name = school["kurum_adi"]

    new_count = 0
    updated_count = 0

    dosya_adi = parsed.get("dosya_adi", "mebbis_raporu.xls")

    for item in parsed.get("items", []):
        barcode = item["barkod_no"]
        c.execute("SELECT id, saha_durumu FROM risk_assessments WHERE school_id = ? AND barkod_no = ?", (school_id, barcode))
        existing = c.fetchone()

        if existing:
            # Güncelle, ancak saha uzmanının belirlediği saha_durumunu koru (eğer daha önce teyit edilmişse)
            current_status = existing["saha_durumu"] or "TEYIT_BEKLIYOR"
            c.execute("""
                UPDATE risk_assessments SET
                    revizyon_no = ?,
                    kontrol_edilen_birim = ?,
                    bolum_unite = ?,
                    tehlike = ?,
                    tehlike_unsurlari = ?,
                    riskler = ?,
                    olasilik = ?,
                    siddet = ?,
                    risk_puani = ?,
                    risk_skoru = ?,
                    risk_seviyesi = ?,
                    sorumlular = ?,
                    alinan_onlemler = ?,
                    alinmasi_gereken_dof_onlem = ?,
                    dayanak_mevzuat = ?,
                    termin_tarihi = ?,
                    mebbis_dosya_adi = ?
                WHERE id = ?
            """, (
                item["revizyon_no"], item["kontrol_edilen_birim"], item["bolum_unite"],
                item["tehlike"], item["tehlike_unsurlari"], item["riskler"],
                item["olasilik"], item["siddet"], item["risk_puani"], item["risk_skoru"],
                item["risk_seviyesi"], item["sorumlular"], item["alinan_onlemler"],
                item["alinmasi_gereken_dof_onlem"], item["dayanak_mevzuat"], item["termin_tarihi"],
                dosya_adi, existing["id"]
            ))
            updated_count += 1
        else:
            # Yeni ekle
            c.execute("""
                INSERT INTO risk_assessments (
                    school_id, barkod_no, revizyon_no, kontrol_edilen_birim, bolum_unite,
                    tehlike, tehlike_unsurlari, riskler, olasilik, siddet,
                    risk_puani, risk_skoru, risk_seviyesi, sorumlular,
                    alinan_onlemler, alinmasi_gereken_dof_onlem, dayanak_mevzuat,
                    termin_tarihi, dof_durumu, kaynak, saha_durumu, denetleyen_uzman,
                    mebbis_dosya_adi
                ) VALUES (
                    ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?,
                    ?, ?, ?, ?,
                    ?, ?, ?,
                    ?, 'Devam Ediyor', 'MEBBIS_RAPORU', 'TEYIT_BEKLIYOR', ?,
                    ?
                )
            """, (
                school_id, barcode, item["revizyon_no"], item["kontrol_edilen_birim"], item["bolum_unite"],
                item["tehlike"], item["tehlike_unsurlari"], item["riskler"], item["olasilik"], item["siddet"],
                item["risk_puani"], item["risk_skoru"], item["risk_seviyesi"], item["sorumlular"],
                item["alinan_onlemler"], item["alinmasi_gereken_dof_onlem"], item["dayanak_mevzuat"],
                item["termin_tarihi"], user_name, dosya_adi
            ))
            new_count += 1

    # İçe aktarım kaydı
    c.execute("""
        INSERT INTO mebbis_imports (school_id, dosya_adi, toplam_madde, yeni_eklenen, guncellenen, yukleyen_kullanici)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (school_id, dosya_adi, len(parsed.get("items", [])), new_count, updated_count, user_name))

    conn.commit()

    return {
        "success": True,
        "school_id": school_id,
        "school_name": school_name,
        "kurum_kodu": school["kurum_kodu"],
        "total_parsed": len(parsed.get("items", [])),
        "new_inserted": new_count,
        "updated": updated_count
    }


def auto_import_local_reports(dir_path: str = r"C:\iş_güvenliği\risk_raporları") -> List[Dict[str, Any]]:
    """Yerel klasördeki tüm MEBBİS raporlarını otomatik olarak eşleştirip veritabanına yükler."""
    if not os.path.exists(dir_path):
        return [{"error": f"Klasör bulunamadı: {dir_path}"}]

    conn = sqlite3.connect("isg_muratpasa.db")
    results = []

    files = sorted([f for f in os.listdir(dir_path) if f.endswith(".xls") or f.endswith(".xlsx")])
    for f in files:
        fp = os.path.join(dir_path, f)
        try:
            if f.endswith(".xls"):
                parsed = parse_mebbis_xls(fp)
            else:
                parsed = parse_mebbis_xlsx(fp)

            res = import_mebbis_parsed_data(conn, parsed)
            results.append(res)
        except Exception as e:
            results.append({"dosya": f, "success": False, "error": str(e)})

    conn.close()
    return results


if __name__ == '__main__':
    print("Yerel 6 MTAL MEBBİS Raporları Otomatik İçe Aktarılıyor...")
    res = auto_import_local_reports()
    for r in res:
        print(r)

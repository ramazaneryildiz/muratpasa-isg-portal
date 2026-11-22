"""
Mevcut Durum Risk Analizi ve DÖF Raporu Üretim Motoru (.docx)
Muratpaşa İlçe MEM İSG Portalı - 2026
"""

import io
import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn


def set_cell_background(cell, hex_color):
    """Hücre arka plan rengini ayarlar."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    """Hücre iç kenar boşluklarını ayarlar."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def add_header(doc, title_text, subtitle_text):
    """Resmî MEB antet başlığı ekler."""
    p_gov = doc.add_paragraph()
    p_gov.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_gov = p_gov.add_run("T.C.\nANTALYA VALİLİĞİ\nMURATPAŞA İLÇE MİLLİ EĞİTİM MÜDÜRLÜĞÜ\nİş Sağlığı ve Güvenliği Bürosu")
    r_gov.font.name = "Calibri"
    r_gov.font.size = Pt(11)
    r_gov.font.bold = True
    r_gov.font.color.rgb = RGBColor(30, 41, 59)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run(title_text)
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(14)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(185, 28, 28)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    r_sub = p_sub.add_run(subtitle_text)
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(10)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(100, 116, 139)


def generate_mevcut_durum_docx(school: dict, risks: list, user_name: str = "Ramazan ERYILDIZ") -> io.BytesIO:
    """Okulun tüm MEBBİS ve Saha tespitlerini içeren 'Mevcut Durum Risk Analizi ve İnceleme Dosyası'nı üretir."""
    doc = Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    school_name = school.get("kurum_adi", "Okul / Kurum")
    kurum_kodu = school.get("kurum_kodu", "")
    today_str = datetime.datetime.now().strftime("%d.%m.%Y")

    add_header(
        doc,
        f"{school_name.upper()}",
        f"MEVCUT DURUM İSG RİSK DEĞERLENDİRME VE İNCELEME DOSYASI (KOD: {kurum_kodu})"
    )

    # 1. Okul Künye Tablosu
    info_table = doc.add_table(rows=5, cols=4)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.autofit = False

    info_data = [
        [("Kurum Adı:", True), (school_name, False), ("Kurum Kodu:", True), (kurum_kodu, False)],
        [("Okul Müdürü (İşveren):", True), (school.get("mudur_ad_soyad") or "Belirtilmedi", False), ("Müdür İletişim:", True), (school.get("mudur_tel") or "Belirtilmedi", False)],
        [("İSG İşveren Vekili:", True), (school.get("isveren_vekili_ad_soyad") or "Belirtilmedi", False), ("Vekil İletişim:", True), (school.get("isveren_vekili_tel") or "Belirtilmedi", False)],
        [("Sorumlu İGU:", True), (user_name, False), ("Tehlike Sınıfı:", True), (school.get("tehlike_sinifi") or "Tehlikeli", False)],
        [("İnceleme / Rapor Tarihi:", True), (today_str, False), ("Dayanak:", True), ("6331 s. Kanun & MEB İSG Yönergesi", False)]
    ]

    for row_idx, row_content in enumerate(info_data):
        for col_idx, (text, is_bold) in enumerate(row_content):
            cell = info_table.cell(row_idx, col_idx)
            set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
            if is_bold:
                set_cell_background(cell, "F1F5F9")
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            run.font.name = "Calibri"
            run.font.size = Pt(9.5)
            run.font.bold = is_bold
            if is_bold:
                run.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph().paragraph_format.space_before = Pt(8)

    # 2. İstatistik & Özet Metrikleri
    tot_count = len(risks)
    giderilen_count = sum(1 for r in risks if r.get("saha_durumu") == "GIDERILDI")
    suren_count = sum(1 for r in risks if r.get("saha_durumu") == "DEVAM_EDIYOR")
    bekleyen_count = sum(1 for r in risks if r.get("saha_durumu") in ("TEYIT_BEKLIYOR", None, ""))
    yeni_count = sum(1 for r in risks if r.get("kaynak") == "SAHA_TESPITI" or r.get("saha_durumu") == "YENI_TESPIT")

    cok_yuksek = sum(1 for r in risks if r.get("risk_seviyesi") == "ÇOK YÜKSEK")
    yuksek = sum(1 for r in risks if r.get("risk_seviyesi") == "YÜKSEK")
    orta = sum(1 for r in risks if r.get("risk_seviyesi") == "ORTA")
    dusuk = sum(1 for r in risks if r.get("risk_seviyesi") == "DÜŞÜK")

    p_metric_title = doc.add_paragraph()
    r_mt = p_metric_title.add_run("1. MEVCUT DURUM İCMAL VE UYGUNLUK ÖZETİ")
    r_mt.font.name = "Calibri"
    r_mt.font.size = Pt(11)
    r_mt.font.bold = True
    r_mt.font.color.rgb = RGBColor(30, 41, 59)
    p_metric_title.paragraph_format.space_after = Pt(4)

    metric_table = doc.add_table(rows=2, cols=5)
    metric_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    metric_headers = ["Toplam Kayıtlı Risk", "Teyit Bekleyen", "Giderilen (Kapalı)", "Devam Eden Risk", "Sahada Yeni Tespit"]
    metric_vals = [str(tot_count), str(bekleyen_count), str(giderilen_count), str(suren_count), str(yeni_count)]
    metric_bg = ["E2E8F0", "FEF08A", "BBF7D0", "FECACA", "FED7AA"]

    for c_idx, h in enumerate(metric_headers):
        cell = metric_table.cell(0, c_idx)
        set_cell_background(cell, "334155")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.name = "Calibri"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    for c_idx, val in enumerate(metric_vals):
        cell = metric_table.cell(1, c_idx)
        set_cell_background(cell, metric_bg[c_idx])
        set_cell_margins(cell, top=120, bottom=120, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(val)
        r.font.name = "Calibri"
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = RGBColor(15, 23, 42)

    doc.add_paragraph().paragraph_format.space_before = Pt(8)

    # 3. Detaylı Risk Listesi Tablosu
    p_list_title = doc.add_paragraph()
    r_lt = p_list_title.add_run("2. BİRİM / MAHAL BAZLI DETAYLI RİSK VE TEHLİKE ENVANTERİ")
    r_lt.font.name = "Calibri"
    r_lt.font.size = Pt(11)
    r_lt.font.bold = True
    r_lt.font.color.rgb = RGBColor(30, 41, 59)
    p_list_title.paragraph_format.space_after = Pt(4)

    cols = ["S.No", "Barkod / No", "Birim / Mahal", "Tehlike & Risk Unsurları", "Puan / Seviye", "Mevcut Saha Durumu", "Alınması Gereken Önlem", "Termin"]
    table = doc.add_table(rows=1, cols=len(cols))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Başlık Satırı
    for c_idx, title in enumerate(cols):
        cell = table.cell(0, c_idx)
        set_cell_background(cell, "1E293B")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.font.name = "Calibri"
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    for idx, rk in enumerate(risks, 1):
        row = table.add_row()
        barcode = rk.get("barkod_no") or f"YENI-{idx:03d}"
        unit = rk.get("kontrol_edilen_birim") or rk.get("bolum_unite") or "-"
        hazard = rk.get("tehlike") or "-"
        elements = rk.get("tehlike_unsurlari") or ""
        hazard_text = f"{hazard}\n({elements})" if elements else hazard

        score = rk.get("risk_puani") or rk.get("risk_skoru") or 0
        level = rk.get("risk_seviyesi") or calculate_risk_level(score)
        status = rk.get("saha_durumu") or "TEYIT_BEKLIYOR"

        status_text_map = {
            "GIDERILDI": ("GİDERİLDİ (KAPALI)", "BBF7D0"),
            "DEVAM_EDIYOR": ("DEVAM EDİYOR (RİSK)", "FECACA"),
            "KISMEN_GIDERILDI": ("KISMEN GİDERİLDİ", "FEF08A"),
            "YENI_TESPIT": ("YENİ SAHA TESPİTİ", "FED7AA"),
            "TEYIT_BEKLIYOR": ("TEYİT BEKLİYOR", "F1F5F9")
        }
        st_text, st_bg = status_text_map.get(status, (status, "FFFFFF"))

        measures = rk.get("alinan_onlemler") or rk.get("alinmasi_gereken_dof_onlem") or "-"
        deadline = rk.get("termin_tarihi") or "-"

        row_data = [
            (str(idx), True, "FFFFFF", WD_ALIGN_PARAGRAPH.CENTER),
            (barcode, False, "FFFFFF", WD_ALIGN_PARAGRAPH.CENTER),
            (unit, True, "FFFFFF", WD_ALIGN_PARAGRAPH.LEFT),
            (hazard_text, False, "FFFFFF", WD_ALIGN_PARAGRAPH.LEFT),
            (f"{score}\n{level}", True, "FFFFFF", WD_ALIGN_PARAGRAPH.CENTER),
            (st_text, True, st_bg, WD_ALIGN_PARAGRAPH.CENTER),
            (measures, False, "FFFFFF", WD_ALIGN_PARAGRAPH.LEFT),
            (deadline, False, "FFFFFF", WD_ALIGN_PARAGRAPH.CENTER),
        ]

        for c_idx, (text, is_bold, bg_color, align) in enumerate(row_data):
            cell = row.cells[c_idx]
            if bg_color != "FFFFFF":
                set_cell_background(cell, bg_color)
            set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
            p = cell.paragraphs[0]
            p.alignment = align
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.name = "Calibri"
            r.font.size = Pt(8)
            r.font.bold = is_bold
            if is_bold and bg_color == "FFFFFF":
                r.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_before = Pt(14)

    # 4. Resmî İmza Bloğu
    sign_table = doc.add_table(rows=2, cols=3)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign_table.autofit = False

    sign_titles = [
        "İşveren Vekili\n(İSG'den Sorumlu Müd. Yrd.)",
        "İş Güvenliği Uzmanı\n(Muratpaşa İlçe MEM)",
        "İşveren\n(Okul Müdürü)"
    ]

    sign_names = [
        f"{school.get('isveren_vekili_ad_soyad') or 'İmza / Kaşe'}",
        f"{user_name}\nİmza",
        f"{school.get('mudur_ad_soyad') or 'İmza / Mühür'}"
    ]

    for c_idx in range(3):
        cell_t = sign_table.cell(0, c_idx)
        set_cell_margins(cell_t, top=60, bottom=60, left=100, right=100)
        p_t = cell_t.paragraphs[0]
        p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_t = p_t.add_run(sign_titles[c_idx])
        r_t.font.name = "Calibri"
        r_t.font.size = Pt(9.5)
        r_t.font.bold = True
        r_t.font.color.rgb = RGBColor(51, 65, 85)

        cell_n = sign_table.cell(1, c_idx)
        set_cell_margins(cell_n, top=200, bottom=60, left=100, right=100)
        p_n = cell_n.paragraphs[0]
        p_n.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_n = p_n.add_run(sign_names[c_idx])
        r_n.font.name = "Calibri"
        r_n.font.size = Pt(9)
        r_n.font.italic = True

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def generate_dof_docx(school: dict, active_risks: list, user_name: str = "Ramazan ERYILDIZ") -> io.BytesIO:
    """Okul idaresine resmi tebliğ edilecek 'Saha Denetim DÖF (Düzeltici ve Önleyici Faaliyet) Raporu'nu üretir."""
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    school_name = school.get("kurum_adi", "Okul / Kurum")
    kurum_kodu = school.get("kurum_kodu", "")
    today_str = datetime.datetime.now().strftime("%d.%m.%Y")
    dof_evrak_no = f"İSG-DÖF-2026/{kurum_kodu}"

    add_header(
        doc,
        f"{school_name.upper()} MÜDÜRLÜĞÜNE",
        "SAHA DENETİMİ DÜZELTİCİ VE ÖNLEYİCİ FAALİYET (DÖF) TEBLİĞ RAPORU"
    )

    # Resmî Evrak Üst Bilgisi
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(4)
    p_meta.paragraph_format.space_after = Pt(8)
    r_m = p_meta.add_run(
        f"Evrak No: {dof_evrak_no}                                             Tarih: {today_str}\n"
        f"Konu: 6331 Sayılı Kanun Kapsamında Tespit Edilen Uygunsuzluklar ve DÖF Eylem Planı"
    )
    r_m.font.name = "Calibri"
    r_m.font.size = Pt(10)
    r_m.font.bold = True
    r_m.font.color.rgb = RGBColor(30, 41, 59)

    # İhtar Paragrafı
    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_after = Pt(8)
    r_d = p_desc.add_run(
        f"İşyerinizde ({school_name}) İş Sağlığı ve Güvenliği Büromuzca gerçekleştirilen yerinde saha denetimlerinde, "
        "önceki MEBBİS Risk Değerlendirmesinde yer alıp halen devam eden tehlikeler ile sahada yeni tespit edilen "
        "uygunsuzluklar aşağıda maddeler halinde çıkarılmıştır. 6331 sayılı İş Sağlığı ve Güvenliği Kanunu'nun 4. ve 11. "
        "maddeleri gereğince belirtilen termin süreleri içerisinde gerekli önlemlerin alınarak sonuçların Büromuza "
        "yazılı ve fotoğraflı olarak bildirilmesi hususunda gereğini rica ederim."
    )
    r_d.font.name = "Calibri"
    r_d.font.size = Pt(9.5)
    r_d.font.color.rgb = RGBColor(71, 85, 105)

    # DÖF Tablosu
    cols = ["DÖF No", "Kontrol Edilen Birim / Mahal", "Tespit Edilen Tehlike & Risk", "Dayanak Mevzuat", "Alınması Gereken DÖF Tedbiri", "Sorumlu", "Kesin Termin", "Durum"]
    table = doc.add_table(rows=1, cols=len(cols))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    for c_idx, title in enumerate(cols):
        cell = table.cell(0, c_idx)
        set_cell_background(cell, "991B1B")  # Kırmızı/DÖF teması
        set_cell_margins(cell, top=100, bottom=100, left=90, right=90)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.font.name = "Calibri"
        r.font.size = Pt(8.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    if not active_risks:
        row = table.add_row()
        cell = row.cells[0]
        cell.merge(row.cells[len(cols) - 1])
        set_cell_margins(cell, top=140, bottom=140, left=140, right=140)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("Tebliğ edilecek açık DÖF veya devam eden uygunsuzluk bulunmamaktadır.")
        r.font.italic = True
        r.font.size = Pt(9.5)
    else:
        for idx, rk in enumerate(active_risks, 1):
            row = table.add_row()
            dof_no = rk.get("dof_no") or f"DÖF-{idx:02d}"
            unit = rk.get("kontrol_edilen_birim") or rk.get("bolum_unite") or "-"
            hazard = rk.get("tehlike") or "-"
            elements = rk.get("tehlike_unsurlari") or ""
            hazard_text = f"{hazard}\n({elements})" if elements else hazard

            legislation = rk.get("dayanak_mevzuat") or "6331 s. Kanun & İlgili Yönetmelikler"
            action = rk.get("alinmasi_gereken_dof_onlem") or rk.get("alinan_onlemler") or "-"
            resp = rk.get("sorumlular") or "Okul İdaresi"
            deadline = rk.get("termin_tarihi") or "15 Gün"
            
            is_new = (rk.get("kaynak") == "SAHA_TESPITI" or rk.get("saha_durumu") == "YENI_TESPIT")
            durum_str = "YENİ TESPİT" if is_new else "DEVAM EDEN RİSK"
            bg_col = "FED7AA" if is_new else "FECACA"

            row_data = [
                (dof_no, True, "FFFFFF", WD_ALIGN_PARAGRAPH.CENTER),
                (unit, True, "FFFFFF", WD_ALIGN_PARAGRAPH.LEFT),
                (hazard_text, False, "FFFFFF", WD_ALIGN_PARAGRAPH.LEFT),
                (legislation, False, "FFFFFF", WD_ALIGN_PARAGRAPH.LEFT),
                (action, True, "FFFFFF", WD_ALIGN_PARAGRAPH.LEFT),
                (resp, False, "FFFFFF", WD_ALIGN_PARAGRAPH.CENTER),
                (deadline, True, "FFFFFF", WD_ALIGN_PARAGRAPH.CENTER),
                (durum_str, True, bg_col, WD_ALIGN_PARAGRAPH.CENTER)
            ]

            for c_idx, (text, is_bold, bg, align) in enumerate(row_data):
                cell = row.cells[c_idx]
                if bg != "FFFFFF":
                    set_cell_background(cell, bg)
                set_cell_margins(cell, top=70, bottom=70, left=70, right=70)
                p = cell.paragraphs[0]
                p.alignment = align
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(text)
                r.font.name = "Calibri"
                r.font.size = Pt(8)
                r.font.bold = is_bold
                if is_bold and bg == "FFFFFF":
                    r.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_before = Pt(16)

    # Tebliğ ve Tebellüğ İmzaları
    sign_table = doc.add_table(rows=2, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign_table.autofit = False

    t_cell = sign_table.cell(0, 0)
    set_cell_margins(t_cell, top=60, bottom=60, left=120, right=120)
    p = t_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("TEBLİĞ EDEN\nİş Güvenliği Uzmanı (İGU)")
    r.font.bold = True
    r.font.size = Pt(10)

    t_cell2 = sign_table.cell(0, 1)
    set_cell_margins(t_cell2, top=60, bottom=60, left=120, right=120)
    p2 = t_cell2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("TEBELLÜĞ EDEN\nİşveren (Okul Müdürü) / İşveren Vekili")
    r2.font.bold = True
    r2.font.size = Pt(10)

    n_cell = sign_table.cell(1, 0)
    set_cell_margins(n_cell, top=200, bottom=60, left=120, right=120)
    p = n_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"{user_name}\nİmza / Tarih: {today_str}")
    r.font.italic = True
    r.font.size = Pt(9.5)

    n_cell2 = sign_table.cell(1, 1)
    set_cell_margins(n_cell2, top=200, bottom=60, left=120, right=120)
    p2 = n_cell2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"{school.get('mudur_ad_soyad') or 'Okul Müdürü'}\nİmza / Mühür / Tarih:")
    r2.font.italic = True
    r2.font.size = Pt(9.5)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


if __name__ == '__main__':
    # Basit test
    test_school = {"kurum_adi": "Antalya Mesleki ve Teknik Anadolu Lisesi", "kurum_kodu": "970149"}
    test_risks = [
        {"barkod_no": "970149BRK000001", "kontrol_edilen_birim": "Kazan Dairesi", "tehlike": "Kazan Patlaması", "saha_durumu": "DEVAM_EDIYOR", "risk_puani": 25, "risk_seviyesi": "ÇOK YÜKSEK"}
    ]
    b1 = generate_mevcut_durum_docx(test_school, test_risks)
    b2 = generate_dof_docx(test_school, test_risks)
    print(f"Mevcut Durum docx boyutu: {len(b1.getvalue())} bytes")
    print(f"DÖF docx boyutu: {len(b2.getvalue())} bytes")

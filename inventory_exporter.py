"""
Okul Fiziki Mekân ve İş Ekipmanı Master Envanter Dosyası Üretici (.xlsx & .docx)
Muratpaşa İlçe MEM İSG Portalı - 2026
"""

import io
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


def generate_inventory_excel(school: dict, meta: dict, spaces: list, items: list, user_name: str = "Ramazan ERYILDIZ") -> io.BytesIO:
    wb = openpyxl.Workbook()

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    sub_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    bold_font = Font(name="Calibri", size=10, bold=True, color="0F172A")
    norm_font = Font(name="Calibri", size=10, color="1E293B")
    risk_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    ok_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    def style_header_row(ws, headers):
        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(1, col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[1].height = 26

    # 1. KURUMSAL VE GÜVENLİK
    ws1 = wb.active
    ws1.title = "1_KURUMSAL_VE_GUVENLIK"
    style_header_row(ws1, ["Alan", "Bilgi"])
    rows_meta = [
        ("Kurum Kodu", school.get("kurum_kodu", "")),
        ("Okul / Kurum Adı", school.get("kurum_adi", "")),
        ("Tehlike Sınıfı", school.get("tehlike_sinifi", "")),
        ("Adres", school.get("adres", "")),
        ("Telefon / E-Posta", f"{school.get('telefon') or '-'} | {school.get('eposta') or '-'}"),
        ("Okul Müdürü (İşveren)", f"{school.get('mudur_ad_soyad') or '-'} ({school.get('mudur_tel') or '-'})"),
        ("İSG Md. Yrd. (İşveren Vekili)", f"{school.get('isveren_vekili_ad_soyad') or '-'} ({school.get('isveren_vekili_tel') or '-'})"),
        ("Çalışan Temsilcisi", meta.get("calisan_temsilcisi", "-")),
        ("Sivil Savunma Amiri", meta.get("sivil_savunma_amiri", "-")),
        ("Doğrulanmış Meslek Alanları", meta.get("dogrulanmis_meslek_alanlari", "-")),
        ("Güvenlik Öncelik / ÖGG / Bahçe Gör.", f"{meta.get('guvenlik_oncelik') or '-'} | ÖGG: {meta.get('ozel_guvenlik_sayisi') or '-'} | Bahçe: {meta.get('bahce_gorevlisi_sayisi') or '-'}"),
        ("Kamera / KGYS / Kayıt Süresi", f"Kamera: {meta.get('kamera_durumu') or '-'} | KGYS: {meta.get('kgys_durumu') or '-'} | Kayıt: {meta.get('kayit_suresi_yeterli') or '-'}"),
        ("Çevre Engeli / Dedektör / Turnike", f"{meta.get('cevre_engeli') or '-'} | El Ded: {meta.get('el_dedektoru') or '-'} | Turnike: {meta.get('turnike') or '-'}"),
        ("Operasyonel Durum Notu", meta.get("operasyonel_not", "-")),
        ("Sorumlu İSG Uzmanı & Tarih", f"{user_name} — {datetime.datetime.now().strftime('%d.%m.%Y')}"),
    ]
    for k, v in rows_meta:
        ws1.append([k, v])
        r_idx = ws1.max_row
        ws1.cell(r_idx, 1).font = bold_font
        ws1.cell(r_idx, 1).fill = sub_fill
        ws1.cell(r_idx, 2).font = norm_font
        ws1.cell(r_idx, 1).border = thin_border
        ws1.cell(r_idx, 2).border = thin_border
    ws1.column_dimensions["A"].width = 34
    ws1.column_dimensions["B"].width = 85

    # 2. FİZİKİ MEKÂN HARİTASI
    ws2 = wb.create_sheet("2_FIZIKI_MEKANLAR")
    style_header_row(ws2, ["ID", "Bina", "Kat", "Mekân / Oda / Atölye Adı", "Mekân Türü", "Ekipman Kalem Sayısı", "Envanter Durumu", "Notlar"])
    space_item_counts = {}
    space_map = {}
    for sp in spaces:
        space_map[sp["id"]] = sp
        space_item_counts[sp["id"]] = 0
    for it in items:
        sid = it.get("space_id")
        if sid in space_item_counts:
            space_item_counts[sid] += 1

    for sp in spaces:
        durum_str = "TAMAMLANDI" if sp.get("durum") == "TAMAMLANDI" else "DEVAM EDİYOR"
        ws2.append([
            sp["id"], sp.get("bina_adi", "Ana Bina"), sp.get("kat", "Genel"),
            sp["mekan_adi"], sp.get("mekan_turu", "Genel Mekân"),
            space_item_counts.get(sp["id"], 0), durum_str, sp.get("notlar") or ""
        ])
        r_idx = ws2.max_row
        for c_i in range(1, 9):
            ws2.cell(r_idx, c_i).font = norm_font
            ws2.cell(r_idx, c_i).border = thin_border
        if sp.get("durum") == "TAMAMLANDI":
            ws2.cell(r_idx, 7).fill = ok_fill

    for col_letter, w in [("A", 8), ("B", 16), ("C", 14), ("D", 38), ("E", 22), ("F", 16), ("G", 18), ("H", 36)]:
        ws2.column_dimensions[col_letter].width = w

    # 3. TÜM EKİPMAN ENVANTERİ
    ws3 = wb.create_sheet("3_EKIPMAN_ENVANTERI")
    style_header_row(ws3, [
        "No", "Bina / Kat / Mekân", "Ekipman / Makine / Demirbaş", "Kategori",
        "Adet", "Marka / Model / Seri", "Çalışma Durumu", "Periyodik Kontrol",
        "Kontrol Etiketi", "İSG Güvenlik Durumu", "İSG Eksik / Risk Detayı", "Kaynak"
    ])
    for idx, it in enumerate(items, 1):
        sp = space_map.get(it.get("space_id"))
        sp_label = f"{sp.get('bina_adi', '')} / {sp.get('kat', '')} / {sp.get('mekan_adi', '')}" if sp else "Okul Geneli / Bağlanmamış"
        pk_str = "EVET (Tabi)" if it.get("periyodik_kontrol_gerekli") else "Hayır"
        ws3.append([
            idx, sp_label, it.get("ekipman_adi", ""), it.get("kategori", ""),
            it.get("adet", 1), it.get("marka_model_seri") or "-", it.get("calisma_durumu", "Aktif"),
            pk_str, it.get("kontrol_etiketi", "-"), it.get("isg_guvenlik_durumu", "Uygun"),
            it.get("isg_eksik_detay") or "-", it.get("kaynak", "")
        ])
        r_idx = ws3.max_row
        for c_i in range(1, 13):
            ws3.cell(r_idx, c_i).font = norm_font
            ws3.cell(r_idx, c_i).border = thin_border
        if it.get("isg_guvenlik_durumu") == "Eksik / Riskli":
            ws3.cell(r_idx, 10).fill = risk_fill
        else:
            ws3.cell(r_idx, 10).fill = ok_fill

    for col_letter, w in [("A", 6), ("B", 34), ("C", 38), ("D", 22), ("E", 8), ("F", 24), ("G", 15), ("H", 16), ("I", 16), ("J", 18), ("K", 36), ("L", 16)]:
        ws3.column_dimensions[col_letter].width = w

    # 4. PERİYODİK KONTROLE TABİ EKİPMANLAR
    ws4 = wb.create_sheet("4_PERIYODIK_KONTROL_TABI")
    style_header_row(ws4, ["No", "Mekân", "Ekipman Adı", "Kategori", "Adet", "Son Kontrol", "Etiket Durumu", "İSG Notu"])
    pk_items = [it for it in items if it.get("periyodik_kontrol_gerekli")]
    for idx, it in enumerate(pk_items, 1):
        sp = space_map.get(it.get("space_id"))
        sp_label = sp.get("mekan_adi") if sp else "Okul Geneli"
        ws4.append([idx, sp_label, it.get("ekipman_adi"), it.get("kategori"), it.get("adet", 1), it.get("son_kontrol_tarihi") or "-", it.get("kontrol_etiketi") or "-", it.get("isg_eksik_detay") or "-"])
    for col_letter, w in [("A", 6), ("B", 30), ("C", 40), ("D", 24), ("E", 8), ("F", 16), ("G", 18), ("H", 36)]:
        ws4.column_dimensions[col_letter].width = w

    # 5. İSG EKSİKLERİ
    ws5 = wb.create_sheet("5_ISG_EKSIKLERI")
    style_header_row(ws5, ["No", "Mekân", "Ekipman Adı", "Adet", "Çalışma Durumu", "Tespit Edilen Eksiklik / Risk"])
    risky_items = [it for it in items if it.get("isg_guvenlik_durumu") == "Eksik / Riskli"]
    for idx, it in enumerate(risky_items, 1):
        sp = space_map.get(it.get("space_id"))
        sp_label = sp.get("mekan_adi") if sp else "Okul Geneli"
        ws5.append([idx, sp_label, it.get("ekipman_adi"), it.get("adet", 1), it.get("calisma_durumu"), it.get("isg_eksik_detay") or "-"])
    for col_letter, w in [("A", 6), ("B", 30), ("C", 36), ("D", 8), ("E", 16), ("F", 50)]:
        ws5.column_dimensions[col_letter].width = w

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def _set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)


def _set_cell_pad(cell, top=80, bottom=80, left=110, right=110):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def generate_inventory_docx(school: dict, meta: dict, spaces: list, items: list, user_name: str = "Ramazan ERYILDIZ") -> io.BytesIO:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.55)
        section.right_margin = Inches(0.55)

    school_name = school.get("kurum_adi", "Okul")
    kurum_kodu = school.get("kurum_kodu", "")
    today_str = datetime.datetime.now().strftime("%d.%m.%Y")

    p_gov = doc.add_paragraph()
    p_gov.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_gov.add_run("T.C.\nMURATPAŞA İLÇE MİLLÎ EĞİTİM MÜDÜRLÜĞÜ\nİş Sağlığı ve Güvenliği Bürosu")
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    r.font.bold = True

    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rt = p_t.add_run(f"{school_name.upper()}\nFİZİKİ MEKÂN VE İŞ EKİPMANI MASTER ENVANTER DOSYASI")
    rt.font.name = "Calibri"
    rt.font.size = Pt(13)
    rt.font.bold = True
    rt.font.color.rgb = RGBColor(30, 58, 138)

    # Özet İstatistikler
    tot_spaces = len(spaces)
    done_spaces = sum(1 for s in spaces if s.get("durum") == "TAMAMLANDI")
    tot_items = len(items)
    pk_count = sum(1 for i in items if i.get("periyodik_kontrol_gerekli"))
    risk_count = sum(1 for i in items if i.get("isg_guvenlik_durumu") == "Eksik / Riskli")

    m_tbl = doc.add_table(rows=2, cols=5)
    m_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Tanımlı Mekân", "Sayımı Tamamlanan", "Toplam Ekipman", "Periyodik Kontrol Tabi", "İSG Eksikli / Riskli"]
    vals = [str(tot_spaces), f"{done_spaces} / {tot_spaces}", str(tot_items), str(pk_count), str(risk_count)]
    bgs = ["E2E8F0", "BBF7D0", "DBEAFE", "FEF08A", "FECACA"]
    for ci, h in enumerate(headers):
        cell = m_tbl.cell(0, ci)
        _set_cell_bg(cell, "1E293B")
        _set_cell_pad(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rh = p.add_run(h)
        rh.font.size = Pt(8.5)
        rh.font.bold = True
        rh.font.color.rgb = RGBColor(255, 255, 255)

        c2 = m_tbl.cell(1, ci)
        _set_cell_bg(c2, bgs[ci])
        _set_cell_pad(c2)
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rv = p2.add_run(vals[ci])
        rv.font.size = Pt(12)
        rv.font.bold = True

    doc.add_paragraph().paragraph_format.space_before = Pt(6)

    # Mekân Bazlı Envanter Tablosu
    p_sec = doc.add_paragraph()
    rs = p_sec.add_run(f"1. KURUM VE MEKÂN BAZLI MAKİNE / TEÇHİZAT / DEMİRBAŞ ENVANTERİ (Tarih: {today_str})")
    rs.font.size = Pt(11)
    rs.font.bold = True

    space_map = {s["id"]: s for s in spaces}
    cols = ["S.No", "Fiziki Mekân", "Ekipman / Demirbaş", "Kategori", "Adet", "Durum", "Periyodik K.", "İSG Güvenlik / Eksik Notu"]
    tbl = doc.add_table(rows=1, cols=len(cols))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ci, h in enumerate(cols):
        cell = tbl.cell(0, ci)
        _set_cell_bg(cell, "334155")
        _set_cell_pad(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rh = p.add_run(h)
        rh.font.size = Pt(8)
        rh.font.bold = True
        rh.font.color.rgb = RGBColor(255, 255, 255)

    for idx, it in enumerate(items, 1):
        row = tbl.add_row()
        sp = space_map.get(it.get("space_id"))
        sp_txt = f"{sp.get('mekan_adi')} ({sp.get('kat', '')})" if sp else "Okul Geneli"
        pk_txt = f"Evet ({it.get('kontrol_etiketi', '')})" if it.get("periyodik_kontrol_gerekli") else "Hayır"
        isg_txt = it.get("isg_eksik_detay") or it.get("isg_guvenlik_durumu", "Uygun")
        vals_row = [
            str(idx), sp_txt, it.get("ekipman_adi", ""), it.get("kategori", ""),
            str(it.get("adet", 1)), it.get("calisma_durumu", "Aktif"), pk_txt, isg_txt
        ]
        for ci, txt in enumerate(vals_row):
            cell = row.cells[ci]
            if ci == 7 and it.get("isg_guvenlik_durumu") == "Eksik / Riskli":
                _set_cell_bg(cell, "FEE2E2")
            _set_cell_pad(cell, top=55, bottom=55, left=70, right=70)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            rc = p.add_run(txt)
            rc.font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_before = Pt(12)
    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    rsign = p_sign.add_run(f"Düzenleyen İş Güvenliği Uzmanı: {user_name}\nKurum Kodu: {kurum_kodu} | Tarih: {today_str}")
    rsign.font.size = Pt(9.5)
    rsign.font.bold = True

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

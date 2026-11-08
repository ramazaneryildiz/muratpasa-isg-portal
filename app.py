import os
import sqlite3
import hashlib
import io
import csv
from datetime import datetime, timedelta
import json
from typing import Optional
import requests
import isg_ai
import isg_storage
from adep_generator import generate_adep_docx, calculate_quotas, get_architectural_units
from okulum_temiz_criteria import OKULUM_TEMIZ_SECTIONS, get_all_questions_flat

from fastapi import FastAPI, Request, Form, Response, Depends, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(BASE_DIR, ".env")
if os.path.exists(env_path):
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            if "=" in line and not line.strip().startswith("#"):
                k, v = line.strip().split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

DB_PATH = os.environ.get("DB_PATH", os.path.join(BASE_DIR, "isg_muratpasa.db"))

app = FastAPI(title="Muratpaşa İlçe MEM İSG Bilgi Yönetim Sistemi")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def hash_pw(password: str) -> str:
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def get_current_user(request: Request):
    user_id = request.cookies.get("session_user_id")
    if not user_id:
        return None
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, full_name, title, role, email FROM users WHERE id = ?", (user_id,))
    user = cursor.fetchone()
    conn.close()
    return user

# ----------------- KULLANICI GİRİŞ & ÇIKIŞ -----------------

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    user = get_current_user(request)
    if user:
        return RedirectResponse("/", status_code=303)
    return templates.TemplateResponse(request=request, name="login.html", context={"user": None, "error": None})

@app.post("/login")
def login_post(request: Request, response: Response, username: str = Form(...), password: str = Form(...)):
    clean_username = username.strip().lower()
    clean_password = password.strip()
    conn = get_db()
    cursor = conn.cursor()
    p_hash = hash_pw(clean_password)
    cursor.execute("SELECT id, username, full_name, role FROM users WHERE LOWER(TRIM(username)) = ? AND password_hash = ?", (clean_username, p_hash))
    user = cursor.fetchone()
    conn.close()

    if not user:
        return templates.TemplateResponse(request=request, name="login.html", context={
            "user": None,
            "error": "Geçersiz kullanıcı adı veya şifre!"
        }, status_code=400)

    res = RedirectResponse("/", status_code=303)
    res.set_cookie(
        key="session_user_id",
        value=str(user["id"]),
        max_age=86400,
        httponly=True,
        samesite="lax"
    )
    return res

@app.get("/logout")
def logout():
    res = RedirectResponse("/login", status_code=303)
    res.delete_cookie("session_user_id")
    return res

# ----------------- İSG VERİ DEĞERLENDİRME & TAMAMLANMA MATRİSİ (20 KRİTER) -----------------

EVALUATION_CRITERIA = [
    ("mudur_ad_soyad", "İletişim & Yönetim", "İşveren (Müdür) Adı-Soyadı", lambda s, m: bool(s.get("mudur_ad_soyad") and s["mudur_ad_soyad"].strip()), "fa-user-tie"),
    ("mudur_tel", "İletişim & Yönetim", "Müdür Cep Telefonu", lambda s, m: bool(s.get("mudur_tel") and s["mudur_tel"].strip()), "fa-phone"),
    ("isveren_vekili_ad_soyad", "İletişim & Yönetim", "İşveren Vekili (İSG Md. Yrd.)", lambda s, m: bool(s.get("isveren_vekili_ad_soyad") and s["isveren_vekili_ad_soyad"].strip()), "fa-user-gear"),
    ("isveren_vekili_tel", "İletişim & Yönetim", "İşveren Vekili Telefonu", lambda s, m: bool(s.get("isveren_vekili_tel") and s["isveren_vekili_tel"].strip()), "fa-phone"),
    ("iletisim_genel", "İletişim & Yönetim", "Kurum E-Posta / Sabit Tel", lambda s, m: bool((s.get("eposta") and s["eposta"].strip()) or (s.get("telefon") and s["telefon"].strip())), "fa-envelope"),

    ("ogrenci_sayisi", "Kapasite & Tesis", "Öğrenci Sayısı", lambda s, m: bool(s.get("ogrenci_sayisi") and s["ogrenci_sayisi"] > 0), "fa-graduation-cap"),
    ("ogretmen_sayisi", "Kapasite & Tesis", "Öğretmen / Personel Sayısı", lambda s, m: bool((s.get("ogretmen_sayisi") and s["ogretmen_sayisi"] > 0) or (s.get("personel_sayisi") and s["personel_sayisi"] > 0)), "fa-chalkboard-user"),
    ("derslik_sayisi", "Kapasite & Tesis", "Derslik Sayısı", lambda s, m: bool(s.get("derslik_sayisi") and s["derslik_sayisi"] > 0), "fa-door-open"),

    ("kamera_sistemi_var_mi", "Güvenlik & Kameralar", "Güvenlik Kamera Durumu", lambda s, m: bool(m and m.get("kamera_sistemi_var_mi")), "fa-video"),
    ("toplam_kamera_sayisi", "Güvenlik & Kameralar", "Toplam Kamera Sayısı", lambda s, m: bool(m and m.get("toplam_kamera_sayisi") and m["toplam_kamera_sayisi"] > 0), "fa-camera"),
    ("kamera_30gun_kayit", "Güvenlik & Kameralar", "30 Gün Kamera Kayıt & Açı", lambda s, m: bool(m and m.get("kamera_30gun_kayit_ve_aci_tam_mi")), "fa-hard-drive"),
    ("ziyaretci_kayit_sistemi", "Güvenlik & Kameralar", "Ziyaretçi Kayıt Defteri/Sistemi", lambda s, m: bool(m and m.get("ziyaretci_kayit_sistemi")), "fa-book-open"),

    ("adep", "Acil Durum & Tatbikat", "ADEP Onay ve Revizyon Tarihi", lambda s, m: bool(m and m.get("adep_var_mi") == 'Var' and m.get("adep_tarihi") and m["adep_tarihi"].strip()), "fa-shield-heart"),
    ("ek7_olcek", "Acil Durum & Tatbikat", "EK-7 Durum Ölçeği Tarihi", lambda s, m: bool(m and m.get("ek7_olcek_var_mi") == 'Var' and m.get("ek7_tarihi") and m["ek7_tarihi"].strip()), "fa-clipboard-check"),
    ("tatbikat_tarihi", "Acil Durum & Tatbikat", "Yangın/Tahliye Tatbikat Tarihi", lambda s, m: bool(m and m.get("tatbikat_tarihi") and m["tatbikat_tarihi"].strip()), "fa-fire-extinguisher"),

    ("kazan_kontrol", "Periyodik Muayeneler", "Kalorifer Kazanı Muayenesi", lambda s, m: bool(m and (m.get("kazan_durumu") == 'Yok' or (m.get("kazan_kontrol_tarihi") and str(m["kazan_kontrol_tarihi"]).strip()))), "fa-temperature-arrow-up"),
    ("elektrik_topraklama", "Periyodik Muayeneler", "Elektrik Topraklama Ölçümü", lambda s, m: bool(m and m.get("elektrik_topraklama_tarihi") and str(m["elektrik_topraklama_tarihi"]).strip()), "fa-bolt"),
    ("paratoner_kontrol", "Periyodik Muayeneler", "Paratoner (Yıldırımlık) Kontrolü", lambda s, m: bool(m and (m.get("paratoner_durumu") == 'Yok' or (m.get("paratoner_kontrol_tarihi") and str(m["paratoner_kontrol_tarihi"]).strip()))), "fa-cloud-bolt"),
    ("yangin_tupleri", "Periyodik Muayeneler", "Yangın Söndürme Tüpleri (YSC)", lambda s, m: bool(m and m.get("yangin_tupleri_kontrol_tarihi") and str(m["yangin_tupleri_kontrol_tarihi"]).strip()), "fa-fire"),

    ("ilkyardimci", "İSG Teşkilatı", "İlkyardımcı Personel Bilgisi", lambda s, m: bool(m and m.get("ilkyardimci_personel_sayisi") and m["ilkyardimci_personel_sayisi"] > 0), "fa-kit-medical"),

    ("sifir_atik", "Çevre & Hijyen", "Sıfır Atık Belgesi", lambda s, m: bool(m and m.get("sifir_atik_belge_no") and str(m["sifir_atik_belge_no"]).strip()), "fa-recycle"),
    ("siginak", "Kapasite & Tesis", "Sığınak Varlık ve Kapasite", lambda s, m: bool(s.get("siginak_var_mi") == "Var"), "fa-person-shelter"),
    ("guvenlik_izni", "Güvenlik & Kameralar", "Özel Güvenlik İzin Belgesi", lambda s, m: bool(m and m.get("guvenlik_izin_belgesi_var_mi") == "VAR"), "fa-id-badge"),
    ("emniyet_puani", "Güvenlik & Kameralar", "Emniyet/Kolluk Öncelik Puanı", lambda s, m: bool(s.get("toplam_guvenlik_puani") and s["toplam_guvenlik_puani"] > 0), "fa-shield-halved"),
]

def evaluate_school_data(school, metrics):
    s = dict(school) if school else {}
    m = dict(metrics) if metrics else {}
    completed_items = []
    missing_items = []
    categories = {}

    for c_id, category, label, check_fn, icon in EVALUATION_CRITERIA:
        if category not in categories:
            categories[category] = {'completed': 0, 'total': 0}
        categories[category]['total'] += 1
        
        is_ok = False
        try:
            is_ok = bool(check_fn(s, m))
        except Exception:
            is_ok = False
            
        item = {
            'id': c_id,
            'category': category,
            'label': label,
            'icon': icon,
            'status': is_ok
        }
        
        if is_ok:
            completed_items.append(item)
            categories[category]['completed'] += 1
        else:
            missing_items.append(item)
            
    total_count = len(EVALUATION_CRITERIA)
    completed_count = len(completed_items)
    percentage = round((completed_count / total_count) * 100) if total_count > 0 else 0
    
    if percentage >= 80:
        color = 'emerald'
        color_class = 'text-emerald-700 bg-emerald-50 border-emerald-200'
        bar_class = 'bg-emerald-500'
    elif percentage >= 50:
        color = 'amber'
        color_class = 'text-amber-700 bg-amber-50 border-amber-200'
        bar_class = 'bg-amber-500'
    else:
        color = 'red'
        color_class = 'text-red-700 bg-red-50 border-red-200'
        bar_class = 'bg-red-500'

    for cat in categories:
        tot = categories[cat]['total']
        comp = categories[cat]['completed']
        categories[cat]['pct'] = round((comp / tot) * 100) if tot > 0 else 0

    return {
        'percentage': percentage,
        'completed_count': completed_count,
        'missing_count': len(missing_items),
        'total_count': total_count,
        'status_color': color,
        'color_class': color_class,
        'bar_class': bar_class,
        'completed_items': completed_items,
        'missing_items': missing_items,
        'categories': categories
    }

# ----------------- DASHBOARD & OKUL LİSTESİ -----------------

@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request, filter: str = "all"):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()

    # İstatistikler (268 Okul Standart İcmali)
    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1")
    total_schools = cursor.fetchone()[0] or 268

    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1 AND (kategori = 'RESMI' OR kurum_statusu = 'Resmî')")
    official_schools = cursor.fetchone()[0] or 122

    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1 AND (kategori = 'OZEL' OR kurum_statusu = 'Özel')")
    private_schools = cursor.fetchone()[0] or 146

    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1 AND (kategori = 'RESMI' OR kurum_statusu = 'Resmî') AND son_veri_giris_tarihi IS NOT NULL")
    completed_resmi = cursor.fetchone()[0] or 95
    if completed_resmi > 95: completed_resmi = 95
    remaining_resmi = official_schools - completed_resmi
    completion_ratio = round((completed_resmi / official_schools) * 100, 1) if official_schools else 77.9

    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1 AND oncelik_derecesi = 1")
    priority_1_count = 11

    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1 AND sorumlu_uzman_id = ?", (user["id"],))
    my_schools_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM risk_assessments WHERE dof_durumu != 'Kapatıldı'")
    open_dof_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM periodic_controls WHERE alarm_seviyesi = 'Suresi Gecmis'")
    expired_controls = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM periodic_controls WHERE alarm_seviyesi = 'Yaklasiyor (30 Gun)'")
    upcoming_controls = cursor.fetchone()[0]

    stats = {
        "total_schools": total_schools,
        "official_schools": official_schools,
        "private_schools": private_schools,
        "completed_resmi": completed_resmi,
        "remaining_resmi": remaining_resmi,
        "completion_ratio": completion_ratio,
        "priority_1_count": priority_1_count,
        "my_schools_count": my_schools_count,
        "open_dof_count": open_dof_count,
        "expired_controls": expired_controls,
        "upcoming_controls": upcoming_controls,
        # Modül 1 İcmal Verileri
        "p1": 11, "p2": 11, "p3": 39, "p4": 34,
        "personel": {
            "ozel_guvenlik": 18,
            "bahce_gorevlisi": 77,
            "sertifikali_bahce": 53,
            "sertifikasiz_bahce": 24,
            "sertifika_orani": 68.8,
            "okul_kolluk": 10,
            "koordinasyon_gorevli": 122,
            "devriye_ekip": 10
        },
        "fiziki": {
            "kamera_var": 95,
            "kamera_oran": 100,
            "kgys_entegre": 7,
            "tek_kapi": 82,
            "cok_kapi": 13,
            "el_dedektoru": 24,
            "kapi_dedektoru": 1,
            "turnike": 12,
            "cevre_engeli": 94,
            "servis_park": 39
        },
        "tedbirler": {
            "danismanlik": 95,
            "cevre_eksik_giderilen": 27,
            "ziyaretci_defteri": 82,
            "yeni_park": 6
        }
    }

    # Filtreye göre okulları listele (Yalnızca 268 aktif okul)
    query = """
    SELECT s.*,
           u.full_name as uzman_adi,
           im.kgys_entegre_mi, im.kamera_sistemi_var_mi, im.toplam_kamera_sayisi, im.kamera_30gun_kayit_ve_aci_tam_mi,
           im.ziyaretci_kayit_sistemi, im.adep_var_mi, im.adep_tarihi, im.ek7_olcek_var_mi, im.ek7_tarihi,
           im.tatbikat_tarihi, im.kazan_kontrol_tarihi, im.elektrik_topraklama_tarihi,
           im.paratoner_kontrol_tarihi, im.yangin_tupleri_kontrol_tarihi,
           im.calisan_temsilcisi_var_mi, im.ilkyardimci_personel_sayisi,
           (SELECT COUNT(*) FROM risk_assessments WHERE school_id = s.id AND dof_durumu != 'Kapatıldı') as open_dof_count
    FROM schools s
    LEFT JOIN users u ON s.sorumlu_uzman_id = u.id
    LEFT JOIN inspection_metrics im ON s.id = im.school_id
    WHERE s.is_active_school = 1
    """
    params = []

    if filter == "my":
        query += " AND s.sorumlu_uzman_id = ?"
        params.append(user["id"])
    elif filter == "official":
        query += " AND (s.kategori = 'RESMI' OR s.kurum_statusu = 'Resmî')"
    elif filter == "private":
        query += " AND (s.kategori = 'OZEL' OR s.kurum_statusu = 'Özel')"
    elif filter == "priority1":
        query += " AND s.oncelik_derecesi = 1"

    query += " ORDER BY s.oncelik_derecesi ASC, s.kurum_adi ASC"
    cursor.execute(query, params)
    raw_schools = cursor.fetchall()

    cursor.execute("""
        SELECT p.*, s.kurum_adi, s.id as school_id
        FROM periodic_controls p
        JOIN schools s ON p.school_id = s.id
        WHERE p.alarm_seviyesi IN ('Suresi Gecmis', 'Yaklasiyor (30 Gun)')
        ORDER BY p.alarm_seviyesi DESC, p.sonraki_kontrol_tarihi ASC
        LIMIT 6
    """)
    urgent_periodic = [dict(r) for r in cursor.fetchall()]

    cursor.execute("""
        SELECT r.*, s.kurum_adi, s.id as school_id
        FROM risk_assessments r
        JOIN schools s ON r.school_id = s.id
        WHERE r.dof_durumu != 'Kapatıldı'
        ORDER BY r.risk_puani DESC
        LIMIT 6
    """)
    urgent_risks = [dict(r) for r in cursor.fetchall()]

    conn.close()

    evaluated_schools = []
    for s in raw_schools:
        s_dict = dict(s)
        comp = evaluate_school_data(s_dict, s_dict)
        s_dict["completion_pct"] = comp["percentage"]
        s_dict["missing_count"] = comp["missing_count"]
        s_dict["color_class"] = comp["color_class"]
        s_dict["bar_class"] = comp["bar_class"]
        evaluated_schools.append(s_dict)

    return templates.TemplateResponse(request=request, name="dashboard.html", context={
        "user": user,
        "stats": stats,
        "schools": evaluated_schools,
        "urgent_periodic": urgent_periodic,
        "urgent_risks": urgent_risks,
        "current_filter": filter,
        "active_page": "dashboard"
    })

@app.get("/api/okul-ara")
def api_search_schools(q: str = ""):
    q = q.strip().lower()
    if not q:
        return []
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, kurum_kodu, kurum_adi, kurum_statusu, kurum_turu, oncelik_derecesi 
        FROM schools 
        WHERE LOWER(kurum_adi) LIKE ? OR kurum_kodu LIKE ?
        ORDER BY oncelik_derecesi ASC, kurum_adi ASC
        LIMIT 10
    """, (f"%{q}%", f"%{q}%"))
    results = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return results

@app.get("/okullar")
def schools_redirect():
    return RedirectResponse("/?filter=all", status_code=303)

# ----------------- OKUL DETAY KARTI -----------------

@app.get("/okul/{school_id}", response_class=HTMLResponse)
def school_detail(request: Request, school_id: int):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()

    # Okul bilgisi
    cursor.execute("""
    SELECT s.*, u.full_name as uzman_adi, u.email as uzman_email, u.title as uzman_unvan, u.igu_sinifi as uzman_sinifi, u.kadro_yeri as uzman_kadro
    FROM schools s
    LEFT JOIN users u ON s.sorumlu_uzman_id = u.id
    WHERE s.id = ?
    """, (school_id,))
    school = cursor.fetchone()
    if not school:
        conn.close()
        raise HTTPException(status_code=404, detail="Kurum bulunamadı.")

    # Metrikler
    cursor.execute("SELECT * FROM inspection_metrics WHERE school_id = ?", (school_id,))
    metrics = cursor.fetchone()

    # Riskler
    cursor.execute("SELECT * FROM risk_assessments WHERE school_id = ? ORDER BY risk_puani DESC", (school_id,))
    risks = cursor.fetchall()

    # Periyodik Kontroller
    cursor.execute("SELECT * FROM periodic_controls WHERE school_id = ? ORDER BY id DESC", (school_id,))
    periodics = cursor.fetchall()

    # İSG Kontrol Listesi (Checklist)
    cursor.execute("SELECT * FROM school_isg_checklists WHERE school_id = ? ORDER BY kategori_no ASC, id ASC", (school_id,))
    checklists = cursor.fetchall()

    # Arşiv Belgeleri
    cursor.execute("SELECT * FROM documents_archive WHERE school_id = ? ORDER BY id DESC", (school_id,))
    documents = cursor.fetchall()

    conn.close()

    completion = evaluate_school_data(dict(school), dict(metrics) if metrics else {})

    return templates.TemplateResponse(request=request, name="school_detail.html", context={
        "user": user,
        "school": school,
        "metrics": metrics,
        "risks": risks,
        "periodics": periodics,
        "checklists": checklists,
        "documents": documents,
        "completion": completion,
        "active_page": "schools"
    })

@app.post("/okul/{school_id}/checklist-guncelle")
async def update_school_checklist(request: Request, school_id: int):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    form_data = await request.form()
    conn = get_db()
    cursor = conn.cursor()

    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
    for key, val in form_data.items():
        if key.startswith("status_"):
            item_id = int(key.split("_")[1])
            status_val = str(val).strip()
            note_val = str(form_data.get(f"note_{item_id}", "")).strip()
            cursor.execute("""
                UPDATE school_isg_checklists
                SET durum = ?, aciklama = ?, denetleyen_uzman_id = ?, denetim_tarihi = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ? AND school_id = ?
            """, (status_val, note_val, user["id"], now_str, item_id, school_id))

    conn.commit()
    conn.close()
    return RedirectResponse(f"/okul/{school_id}?msg=checklist_ok#checklistTab", status_code=303)

@app.get("/okul/{school_id}/checklist-yazdir", response_class=HTMLResponse)
def print_school_checklist(request: Request, school_id: int):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT s.*, u.full_name as uzman_adi, u.igu_sinifi as uzman_sinifi
        FROM schools s
        LEFT JOIN users u ON s.sorumlu_uzman_id = u.id
        WHERE s.id = ?
    """, (school_id,))
    school = cursor.fetchone()
    if not school:
        conn.close()
        raise HTTPException(status_code=404, detail="Kurum bulunamadı.")

    cursor.execute("SELECT * FROM school_isg_checklists WHERE school_id = ? ORDER BY kategori_no ASC, id ASC", (school_id,))
    checklists = cursor.fetchall()
    conn.close()

    categories = {}
    for c in checklists:
        k_adi = c["kategori_adi"]
        if k_adi not in categories:
            categories[k_adi] = {"no": c["kategori_no"], "maddeler": []}
        categories[k_adi]["maddeler"].append(c)

    return templates.TemplateResponse(request=request, name="checklist_print.html", context={
        "user": user,
        "school": school,
        "categories": categories,
        "now_date": datetime.now().strftime("%d.%m.%Y")
    })

@app.post("/okul/{school_id}/guncelle")
def update_school(
    school_id: int,
    mudur_ad_soyad: str = Form(""),
    mudur_tel: str = Form(""),
    isveren_vekili_ad_soyad: str = Form(""),
    isveren_vekili_tel: str = Form(""),
    oncelik_derecesi: int = Form(3),
    kgys_entegre_mi: str = Form("Hayır")
):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
    UPDATE schools SET
        mudur_ad_soyad = ?, mudur_tel = ?,
        isveren_vekili_ad_soyad = ?, isveren_vekili_tel = ?,
        oncelik_derecesi = ?
    WHERE id = ?
    """, (mudur_ad_soyad, mudur_tel, isveren_vekili_ad_soyad, isveren_vekili_tel, oncelik_derecesi, school_id))

    cursor.execute("""
    UPDATE inspection_metrics SET kgys_entegre_mi = ? WHERE school_id = ?
    """, (kgys_entegre_mi, school_id))

    conn.commit()
    conn.close()
    return RedirectResponse(f"/okul/{school_id}", status_code=303)

@app.post("/okul/{school_id}/risk-ekle")
def add_risk(
    school_id: int,
    kontrol_edilen_birim: str = Form(...),
    tehlike: str = Form(...),
    risk_puani: int = Form(10),
    termin_tarihi: str = Form(""),
    alinan_onlemler: str = Form(""),
    dof_durumu: str = Form("Süresi Devam Ediyor")
):
    conn = get_db()
    cursor = conn.cursor()
    barkod_no = f"BRK{school_id}{int(datetime.now().timestamp()) % 10000:04d}"
    cursor.execute("""
    INSERT INTO risk_assessments (
        school_id, barkod_no, revizyon_no, kontrol_edilen_birim, tehlike,
        tehlike_unsurlari, riskler, olasilik, siddet, risk_puani,
        sorumlular, alinan_onlemler, termin_tarihi, dof_durumu
    ) VALUES (?, ?, '1', ?, ?, ?, 'Güvenlik Riski', '2. Düşük', '3. Orta', ?, 'İŞVEREN VEKİLİ', ?, ?, ?)
    """, (school_id, barkod_no, kontrol_edilen_birim, tehlike, tehlike, risk_puani, alinan_onlemler, termin_tarihi, dof_durumu))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/okul/{school_id}", status_code=303)

@app.post("/okul/{school_id}/muayene-ekle")
def add_periodic(
    school_id: int,
    ekipman_turu: str = Form(...),
    son_kontrol_tarihi: str = Form(...),
    sonraki_kontrol_tarihi: str = Form(...),
    kontrol_yapan_kurum: str = Form(...),
    rapor_sonucu: str = Form("Kullanıma Uygun")
):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
    INSERT INTO periodic_controls (
        school_id, ekipman_turu, son_kontrol_tarihi, gecerlilik_ayi,
        sonraki_kontrol_tarihi, kontrol_yapan_kurum, rapor_sonucu,
        alarm_seviyesi, notlar
    ) VALUES (?, ?, ?, 12, ?, ?, ?, 'Gecerli', 'Sistem üzerinden manuel eklendi.')
    """, (school_id, ekipman_turu, son_kontrol_tarihi, sonraki_kontrol_tarihi, kontrol_yapan_kurum, rapor_sonucu))
    conn.commit()
    conn.close()
    return RedirectResponse(f"/okul/{school_id}", status_code=303)

# ----------------- MÜFETTİŞ İCMAL RAPORU (16 TABLO) -----------------

@app.get("/mufettis-raporu", response_class=HTMLResponse)
def inspection_summary(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1")
    total = cursor.fetchone()[0] or 268

    # Tablo 1
    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1 AND oncelik_derecesi = 1")
    p1 = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1 AND oncelik_derecesi = 2")
    p2 = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1 AND oncelik_derecesi = 3")
    p3 = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM schools WHERE is_active_school = 1 AND oncelik_derecesi = 4")
    p4 = cursor.fetchone()[0]

    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND im.adep_var_mi = 'Var'
    """)
    adep_var = cursor.fetchone()[0]
    adep_yok = total - adep_var

    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND im.ek7_olcek_var_mi = 'Var'
    """)
    ek7_var = cursor.fetchone()[0]

    t1 = {
        "total": total,
        "p1": p1, "p2": p2, "p3": p3, "p4": p4,
        "adep_var": adep_var, "adep_yok": adep_yok,
        "adep_ratio": int((adep_var / total) * 100) if total else 0,
        "ek7_var": ek7_var,
        "ek7_ratio": int((ek7_var / total) * 100) if total else 0
    }

    # Tablo 2
    cursor.execute("""
        SELECT SUM(im.ozel_guvenlik_sayisi), SUM(im.bahce_gorevlisi_sayisi), SUM(im.sertifikali_bahce_gorevlisi_sayisi), SUM(im.metal_dedektor_egitimli_sayisi)
        FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1
    """)
    t2_row = cursor.fetchone()
    t2 = {
        "ozel_guvenlik": t2_row[0] or 18,
        "bahce_gorevlisi": t2_row[1] or 77,
        "sertifikali": t2_row[2] or 53,
        "kolluk": 10,
        "dedektor_egitimli": t2_row[3] or 24
    }

    # Tablo 5: Okul Gözlem Raporu İcmali
    cursor.execute("""
        SELECT COALESCE(SUM(im.okul_gozlem_raporu_sayisi), 0)
        FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1
    """)
    gozlem_toplam = cursor.fetchone()[0]
    t5 = {
        "total_schools": total,
        "toplam_gozlem": gozlem_toplam,
        "okul_basi_gozlem": round(gozlem_toplam / total, 2) if total else 0
    }

    # Tablo 6
    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND im.kamera_sistemi_var_mi = 'Var'
    """)
    kamera_var = cursor.fetchone()[0]
    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND im.kgys_entegre_mi = 'Evet'
    """)
    kgys_var = cursor.fetchone()[0]
    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND im.giris_cikis_kapi_durumu LIKE '%Tek Kapı%'
    """)
    tek_kapi = cursor.fetchone()[0]
    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND im.turnike_var_mi = 'Var'
    """)
    turnike_var = cursor.fetchone()[0]
    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND im.fiziki_engel_turu != 'Fiziki Engel Yok'
    """)
    fiziki_var = cursor.fetchone()[0]
    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND im.arac_servis_park_yeri_var_mi = 'Var'
    """)
    park_var = cursor.fetchone()[0]

    t6 = {
        "kamera_var": kamera_var,
        "kamera_ratio": int((kamera_var / total) * 100) if total else 0,
        "kgys_var": kgys_var,
        "kgys_ratio": int((kgys_var / total) * 100) if total else 0,
        "tek_kapi": tek_kapi,
        "turnike_var": turnike_var,
        "fiziki_var": fiziki_var,
        "park_var": park_var
    }

    # Tablo 7: Ziyaretçi Kayıt ve Çevre Düzenlemesi
    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND im.ziyaretci_kayit_sistemi = 'Evet'
    """)
    ziyaretci_var = cursor.fetchone()[0]
    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND im.ozel_arac_bahce_engeli = 'Evet'
    """)
    arac_engel_var = cursor.fetchone()[0]
    cursor.execute("""
        SELECT COUNT(*) FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1 AND (im.el_dedektoru_var_mi = 'Var' OR im.kapi_dedektoru_var_mi = 'Var')
    """)
    dedektor_var = cursor.fetchone()[0]
    t7 = {
        "ziyaretci_var": ziyaretci_var,
        "ziyaretci_ratio": int((ziyaretci_var / total) * 100) if total else 0,
        "arac_engel_var": arac_engel_var,
        "dedektor_var": dedektor_var
    }

    # Tablo 11: Okul Güvenliği Eğitim İcmali (6 Hedef Grup)
    cursor.execute("""
        SELECT 
            COALESCE(SUM(im.egitim_yonetici_toplam), 0), COALESCE(SUM(im.egitim_yonetici_alan), 0),
            COALESCE(SUM(im.egitim_ogretmen_toplam), 0), COALESCE(SUM(im.egitim_ogretmen_alan), 0),
            COALESCE(SUM(im.egitim_pdr_toplam), 0), COALESCE(SUM(im.egitim_pdr_alan), 0),
            COALESCE(SUM(im.egitim_ogrenci_toplam), 0), COALESCE(SUM(im.egitim_ogrenci_alan), 0),
            COALESCE(SUM(im.egitim_veli_toplam), 0), COALESCE(SUM(im.egitim_veli_alan), 0),
            COALESCE(SUM(im.egitim_diger_toplam), 0), COALESCE(SUM(im.egitim_diger_alan), 0)
        FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1
    """)
    e_row = cursor.fetchone()
    t11 = [
        {"grup": "İl / İlçe Yöneticileri", "hedef": 7, "alan": 7, "oran": 100},
        {"grup": "Okul / Kurum Yöneticileri", "hedef": 111, "alan": 111, "oran": 100},
        {"grup": "Öğretmenler", "hedef": e_row[2] or 5420, "alan": e_row[3] or 5180, "oran": round(((e_row[3] or 5180) / (e_row[2] or 5420)) * 100)},
        {"grup": "Rehberlik & Psikolojik Danışman (PDR)", "hedef": e_row[4] or 184, "alan": e_row[5] or 184, "oran": 100},
        {"grup": "Öğrenciler", "hedef": e_row[6] or 84500, "alan": e_row[7] or 79200, "oran": round(((e_row[7] or 79200) / (e_row[6] or 84500)) * 100)},
        {"grup": "Veliler", "hedef": e_row[8] or 45000, "alan": e_row[9] or 32400, "oran": round(((e_row[9] or 32400) / (e_row[8] or 45000)) * 100)},
    ]

    # Tablo 14 & 15: Güvenlik Yönetim Ekibi ve Olaylar
    cursor.execute("""
        SELECT COALESCE(SUM(im.guvenlik_ekibi_toplanti_sayisi), 0),
               COALESCE(SUM(im.guvenlik_olay_sayisi), 0),
               COALESCE(SUM(im.ihmal_istismar_sayisi), 0),
               COALESCE(SUM(im.bildirilen_ihmal_istismar_sayisi), 0)
        FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1
    """)
    m_row = cursor.fetchone()
    t14 = {
        "toplanti_sayisi": m_row[0] or 0,
        "olay_sayisi": m_row[1] or 0,
        "ihmal_istismar": m_row[2] or 0,
        "bildirilen_ihmal": m_row[3] or 0
    }

    # Özel Sektör & Ekstra Modül İcmalleri
    cursor.execute("""
        SELECT 
            COUNT(CASE WHEN siginak_var_mi = 'Var' THEN 1 END),
            COALESCE(SUM(siginak_kapasite_kisi), 0),
            COALESCE(SUM(siginak_alan_m2), 0)
        FROM schools
        WHERE is_active_school = 1
    """)
    sig_row = cursor.fetchone()
    siginak_ozet = {
        "okul_sayisi": sig_row[0],
        "kapasite_kisi": sig_row[1],
        "alan_m2": sig_row[2]
    }

    cursor.execute("""
        SELECT 
            COUNT(CASE WHEN im.sifir_atik_belge_no IS NOT NULL AND im.sifir_atik_belge_no != '' THEN 1 END),
            COALESCE(SUM(im.sifir_atik_egitim_sayisi), 0),
            COALESCE(SUM(im.sifir_atik_katilimci_sayisi), 0)
        FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1
    """)
    sa_row = cursor.fetchone()
    sifir_atik_ozet = {
        "belgeli_okul": sa_row[0],
        "egitim_sayisi": sa_row[1],
        "katilimci_sayisi": sa_row[2]
    }

    cursor.execute("""
        SELECT 
            COALESCE(SUM(im.toplam_kamera_sayisi), 0),
            COALESCE(SUM(im.kamera_bina_ici_calismayan) + SUM(im.kamera_bina_disi_calismayan), 0),
            COALESCE(SUM(im.kamera_ihtiyac_sayisi), 0),
            COUNT(CASE WHEN im.guvenlik_izin_belgesi_var_mi = 'VAR' THEN 1 END)
        FROM inspection_metrics im
        JOIN schools s ON im.school_id = s.id
        WHERE s.is_active_school = 1
    """)
    kam_row = cursor.fetchone()
    kamera_ozet = {
        "toplam_montajli": kam_row[0],
        "toplam_arizali": kam_row[1],
        "acil_ihtiyac": kam_row[2],
        "izin_belgeli": kam_row[3]
    }

    conn.close()

    return templates.TemplateResponse(request=request, name="inspection_summary.html", context={
        "user": user,
        "t1": t1,
        "t2": t2,
        "t5": t5,
        "t6": t6,
        "t7": t7,
        "t11": t11,
        "t14": t14,
        "siginak_ozet": siginak_ozet,
        "sifir_atik_ozet": sifir_atik_ozet,
        "kamera_ozet": kamera_ozet,
        "active_page": "inspection"
    })

# ----------------- ÖZEL GELİŞMİŞ RAPOR SAYFALARI -----------------

@app.get("/raporlar/guvenlik-kamera", response_class=HTMLResponse)
def report_camera(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT s.id, s.kurum_kodu, s.kurum_adi, s.oncelik_derecesi,
               m.toplam_kamera_sayisi, m.kamera_bina_ici_calisan, m.kamera_bina_ici_calismayan,
               m.kamera_bina_disi_calisan, m.kamera_bina_disi_calismayan,
               m.kamera_kayit_gun_sayisi, m.kamera_ihtiyac_sayisi,
               m.kgys_entegre_mi, m.guvenlik_izin_belgesi_var_mi, m.ozel_guvenlik_sayisi, m.guvenlik_istihdam_turu
        FROM schools s
        LEFT JOIN inspection_metrics m ON s.id = m.school_id
        ORDER BY m.toplam_kamera_sayisi DESC, s.kurum_adi ASC
    """)
    schools = cursor.fetchall()

    cursor.execute("""
        SELECT 
            COALESCE(SUM(toplam_kamera_sayisi), 0) as toplam_montajli,
            COALESCE(SUM(kamera_bina_ici_calisan), 0) as bina_ici_calisan,
            COALESCE(SUM(kamera_bina_ici_calismayan), 0) as bina_ici_arizali,
            COALESCE(SUM(kamera_bina_disi_calisan), 0) as bina_disi_calisan,
            COALESCE(SUM(kamera_bina_disi_calismayan), 0) as bina_disi_arizali,
            COALESCE(SUM(kamera_ihtiyac_sayisi), 0) as toplam_ihtiyac,
            COUNT(CASE WHEN kamera_sistemi_var_mi = 'Var' THEN 1 END) as kamera_olan_okul,
            COUNT(CASE WHEN kgys_entegre_mi = 'Evet' THEN 1 END) as kgys_okul,
            COUNT(CASE WHEN guvenlik_izin_belgesi_var_mi = 'VAR' THEN 1 END) as izin_belgesi_var
        FROM inspection_metrics
    """)
    st = cursor.fetchone()
    stats = {
        "toplam_montajli": st["toplam_montajli"],
        "toplam_calisan": st["bina_ici_calisan"] + st["bina_disi_calisan"],
        "toplam_arizali": st["bina_ici_arizali"] + st["bina_disi_arizali"],
        "bina_ici_calisan": st["bina_ici_calisan"],
        "bina_ici_arizali": st["bina_ici_arizali"],
        "bina_disi_calisan": st["bina_disi_calisan"],
        "bina_disi_arizali": st["bina_disi_arizali"],
        "toplam_ihtiyac": st["toplam_ihtiyac"],
        "kamera_olan_okul": st["kamera_olan_okul"],
        "kgys_okul": st["kgys_okul"],
        "izin_belgesi_var": st["izin_belgesi_var"]
    }
    conn.close()

    return templates.TemplateResponse(request=request, name="report_camera.html", context={
        "user": user,
        "schools": schools,
        "stats": stats,
        "active_page": "reports"
    })

@app.get("/raporlar/siginak-afet", response_class=HTMLResponse)
def report_shelter(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, kurum_kodu, kurum_adi, kurum_statusu, adres,
               siginak_var_mi, siginak_kapasite_kisi, siginak_alan_m2
        FROM schools
        ORDER BY siginak_kapasite_kisi DESC, kurum_adi ASC
    """)
    schools = cursor.fetchall()

    cursor.execute("""
        SELECT 
            COUNT(*) as total_schools,
            COUNT(CASE WHEN siginak_var_mi = 'Var' THEN 1 END) as shelter_schools,
            COALESCE(SUM(siginak_kapasite_kisi), 0) as total_capacity_persons,
            COALESCE(SUM(siginak_alan_m2), 0) as total_area_m2
        FROM schools
    """)
    st = cursor.fetchone()
    total_cnt = st["total_schools"] or 1
    shelter_cnt = st["shelter_schools"] or 0
    stats = {
        "total_schools": st["total_schools"],
        "shelter_schools": shelter_cnt,
        "shelter_ratio": round((shelter_cnt / total_cnt) * 100),
        "total_capacity_persons": st["total_capacity_persons"],
        "total_area_m2": st["total_area_m2"]
    }
    conn.close()

    return templates.TemplateResponse(request=request, name="report_shelter.html", context={
        "user": user,
        "schools": schools,
        "stats": stats,
        "active_page": "reports"
    })

@app.get("/raporlar/sifir-atik", response_class=HTMLResponse)
def report_zero_waste(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT s.id, s.kurum_kodu, s.kurum_adi, s.kurum_turu,
               m.sifir_atik_belge_no, m.sifir_atik_belge_tarihi, m.sifir_atik_gecerlilik_tarihi,
               m.sifir_atik_egitim_sayisi, m.sifir_atik_katilimci_sayisi
        FROM schools s
        LEFT JOIN inspection_metrics m ON s.id = m.school_id
        ORDER BY CASE WHEN m.sifir_atik_belge_no IS NOT NULL AND m.sifir_atik_belge_no != '' THEN 0 ELSE 1 END,
                 s.kurum_adi ASC
    """)
    schools = cursor.fetchall()

    cursor.execute("""
        SELECT 
            COUNT(*) as total_schools,
            COUNT(CASE WHEN sifir_atik_belge_no IS NOT NULL AND sifir_atik_belge_no != '' THEN 1 END) as certified_schools,
            COALESCE(SUM(sifir_atik_egitim_sayisi), 0) as total_trainings,
            COALESCE(SUM(sifir_atik_katilimci_sayisi), 0) as total_participants
        FROM inspection_metrics
    """)
    st = cursor.fetchone()
    total_cnt = st["total_schools"] or 1
    cert_cnt = st["certified_schools"] or 0
    stats = {
        "certified_schools": cert_cnt,
        "uncertified_schools": total_cnt - cert_cnt,
        "certified_ratio": round((cert_cnt / total_cnt) * 100),
        "total_trainings": st["total_trainings"],
        "total_participants": st["total_participants"]
    }
    conn.close()

    return templates.TemplateResponse(request=request, name="report_zero_waste.html", context={
        "user": user,
        "schools": schools,
        "stats": stats,
        "active_page": "reports"
    })

# ----------------- RİSK & ALARMLAR SAYFASI -----------------

@app.get("/risk-alarmlari", response_class=HTMLResponse)
def risk_alarms(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT r.*, s.kurum_adi
    FROM risk_assessments r
    JOIN schools s ON r.school_id = s.id
    WHERE r.dof_durumu LIKE '%SÜRESİ GEÇMİŞ%'
    ORDER BY r.risk_puani DESC
    """)
    open_dofs = cursor.fetchall()

    cursor.execute("""
    SELECT p.*, s.kurum_adi
    FROM periodic_controls p
    JOIN schools s ON p.school_id = s.id
    WHERE p.alarm_seviyesi IN ('Suresi Gecmis', 'Yaklasiyor (30 Gun)')
    ORDER BY p.alarm_seviyesi DESC, p.sonraki_kontrol_tarihi ASC
    """)
    periodic_alarms = cursor.fetchall()

    conn.close()

    return templates.TemplateResponse(request=request, name="risk_alarms.html", context={
        "user": user,
        "open_dofs": open_dofs,
        "periodic_alarms": periodic_alarms,
        "active_page": "risks"
    })

@app.get("/export/periyodik-muayeneler")
def export_periodic_controls(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT s.kurum_kodu, s.kurum_adi, s.kurum_statusu,
               p.ekipman_turu, p.son_kontrol_tarihi, p.sonraki_kontrol_tarihi,
               p.kontrol_yapan_kurum, p.rapor_sonucu, p.alarm_seviyesi, p.notlar
        FROM periodic_controls p
        JOIN schools s ON p.school_id = s.id
        ORDER BY 
            CASE p.alarm_seviyesi 
                WHEN 'Suresi Gecmis' THEN 1 
                WHEN 'Yaklasiyor (30 Gun)' THEN 2 
                ELSE 3 
            END,
            p.sonraki_kontrol_tarihi ASC
    """)
    rows = cursor.fetchall()
    conn.close()

    output = io.StringIO()
    output.write('\ufeff')
    writer = csv.writer(output, delimiter=';')
    writer.writerow([
        "Kurum Kodu", "Kurum Adı", "Statü", "Cihaz / Ekipman Türü",
        "Son Muayene Tarihi", "Sonraki Muayene Tarihi", "Kontrol Yapan Kurum",
        "Rapor Sonucu", "Alarm Seviyesi", "Notlar / Açıklama"
    ])
    for r in rows:
        alarm_text = "SÜRESİ GEÇMİŞ (KRİTİK)" if r["alarm_seviyesi"] == "Suresi Gecmis" else ("YAKLAŞIYOR (30 GÜN)" if r["alarm_seviyesi"] == "Yaklasiyor (30 Gun)" else "GEÇERLİ")
        writer.writerow([
            r["kurum_kodu"], r["kurum_adi"], r["kurum_statusu"], r["ekipman_turu"],
            r["son_kontrol_tarihi"], r["sonraki_kontrol_tarihi"], r["kontrol_yapan_kurum"] or "",
            r["rapor_sonucu"] or "", alarm_text, r["notlar"] or ""
        ])

    today_str = datetime.now().strftime("%Y%m%d")
    filename = f"Muratpasa_Hayati_Periyodik_Muayene_Takvimi_{today_str}.csv"
    return Response(
        content=output.getvalue(),
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# ----------------- EXCEL EXPORT (CSV / UTF-8 BOM) -----------------

@app.get("/export/excel")
def export_excel(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT s.kurum_kodu, s.kurum_adi, s.kurum_statusu, s.bagli_birim, s.kurum_turu,
           s.oncelik_derecesi, s.tehlike_sinifi, s.mudur_ad_soyad, s.mudur_tel,
           s.isveren_vekili_ad_soyad, s.isveren_vekili_tel, u.full_name as sorumlu_uzman,
           im.adep_var_mi, im.adep_tarihi, im.ek7_olcek_var_mi, im.ek7_tarihi,
           im.ozel_guvenlik_sayisi, im.bahce_gorevlisi_sayisi, im.sertifikali_bahce_gorevlisi_sayisi,
           im.kamera_sistemi_var_mi, im.kgys_entegre_mi, im.giris_cikis_kapi_durumu,
           im.turnike_var_mi, im.fiziki_engel_turu, im.arac_servis_park_yeri_var_mi
    FROM schools s
    LEFT JOIN users u ON s.sorumlu_uzman_id = u.id
    LEFT JOIN inspection_metrics im ON s.id = im.school_id
    WHERE s.is_active_school = 1
    ORDER BY s.oncelik_derecesi ASC, s.kurum_adi ASC
    """)
    rows = cursor.fetchall()
    conn.close()

    output = io.StringIO()
    output.write('\ufeff')
    writer = csv.writer(output, delimiter=';')

    writer.writerow([
        "Kurum Kodu", "Kurum Adı", "Statü", "Bağlı Birim", "Kurum Türü",
        "Öncelik Derecesi", "Tehlike Sınıfı", "Okul Müdürü", "Müdür Telefon",
        "İşveren Vekili (İSG Md. Yrd.)", "İşveren Vekili Telefon", "Sorumlu İSG Uzmanı",
        "Acil Durum Planı (ADEP)", "ADEP Tarihi", "EK-7 Değerlendirme Ölçeği", "EK-7 Tarihi",
        "Özel Güvenlik Sayısı", "Bahçe Görevlisi (TYP)", "Sertifikalı Bahçe Görevlisi",
        "Kamera Sistemi", "KGYS Entegrasyonu", "Giriş-Çıkış Kapı",
        "Turnike", "Fiziki Engel", "Servis/Araç Park Alanı"
    ])

    for r in rows:
        writer.writerow(list(r))

    output.seek(0)
    filename = f"Muratpasa_268_Okul_ISG_Guvenlik_Icmali_{datetime.now().strftime('%Y%m%d')}.csv"
    return StreamingResponse(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# ----------------- OKUL VERİ GİRİŞ PORTALI (SELF-SERVICE) -----------------

def get_current_school(request: Request):
    school_id = request.cookies.get("session_school_id")
    if not school_id:
        return None
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM schools WHERE id = ?", (school_id,))
    school = cursor.fetchone()
    conn.close()
    return school

@app.get("/okul-portal", response_class=HTMLResponse)
def school_portal_login_page(request: Request):
    school = get_current_school(request)
    if school:
        return RedirectResponse("/okul-portal/form", status_code=303)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT kurum_kodu, kurum_adi FROM schools ORDER BY kurum_adi ASC")
    schools_list = cursor.fetchall()
    conn.close()

    return templates.TemplateResponse(request=request, name="school_portal_login.html", context={
        "error": None,
        "schools_list": schools_list
    })

@app.post("/okul-portal/login")
def school_portal_login_post(request: Request, response: Response, kurum_kodu: str = Form(...), sifre: Optional[str] = Form(None)):
    clean_code = kurum_kodu.strip().split(' ')[0].split('-')[0].strip()
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, kurum_kodu, portal_sifre, kurum_adi FROM schools WHERE kurum_kodu = ?", (clean_code,))
    school = cursor.fetchone()

    if not school:
        cursor.execute("SELECT id, kurum_kodu, portal_sifre, kurum_adi FROM schools WHERE kurum_adi LIKE ?", (f"%{kurum_kodu.strip()}%",))
        school = cursor.fetchone()

    if not school:
        cursor.execute("SELECT kurum_kodu, kurum_adi FROM schools ORDER BY kurum_adi ASC")
        schools_list = cursor.fetchall()
        conn.close()
        return templates.TemplateResponse(request=request, name="school_portal_login.html", context={
            "error": f"'{kurum_kodu}' ile eşleşen kurum bulunamadı! Lütfen 6 haneli MEB kurum kodunuzu kontrol ediniz.",
            "schools_list": schools_list
        }, status_code=400)

    # Şifre kontrolü (isteğe bağlı güvenlik)
    if sifre and sifre.strip() and sifre.strip() != "isg2026":
        actual_pw = school["portal_sifre"] or "isg2026"
        if sifre.strip() != actual_pw:
            cursor.execute("SELECT kurum_kodu, kurum_adi FROM schools ORDER BY kurum_adi ASC")
            schools_list = cursor.fetchall()
            conn.close()
            return templates.TemplateResponse(request=request, name="school_portal_login.html", context={
                "error": "Kurum şifresi hatalı! (Varsayılan şifre: isg2026)",
                "schools_list": schools_list
            }, status_code=400)

    conn.close()
    res = RedirectResponse("/okul-portal/form", status_code=303)
    res.set_cookie(key="session_school_id", value=str(school["id"]), max_age=86400, httponly=True)
    return res

@app.get("/okul-portal/giris/{access_token}")
def school_portal_direct_access(access_token: str):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM schools WHERE access_token = ?", (access_token.strip(),))
    school = cursor.fetchone()
    conn.close()

    if not school:
        return RedirectResponse("/okul-portal", status_code=303)

    res = RedirectResponse("/okul-portal/form", status_code=303)
    res.set_cookie(key="session_school_id", value=str(school["id"]), max_age=86400, httponly=True)
    return res

@app.get("/okul-portal/form", response_class=HTMLResponse)
def school_portal_form_page(request: Request, saved: int = 0):
    school = get_current_school(request)
    if not school:
        return RedirectResponse("/okul-portal", status_code=303)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM inspection_metrics WHERE school_id = ?", (school["id"],))
    metrics = cursor.fetchone()
    if not metrics:
        cursor.execute("INSERT INTO inspection_metrics (school_id) VALUES (?)", (school["id"],))
        conn.commit()
        cursor.execute("SELECT * FROM inspection_metrics WHERE school_id = ?", (school["id"],))
        metrics = cursor.fetchone()

    conn.close()

    completion = evaluate_school_data(dict(school), dict(metrics) if metrics else {})

    return templates.TemplateResponse(request=request, name="school_portal.html", context={
        "school": school,
        "metrics": metrics,
        "completion": completion,
        "msg": "saved" if saved == 1 else None
    })

@app.get("/okul-portal/yazdir", response_class=HTMLResponse)
def school_portal_print(request: Request):
    school = get_current_school(request)
    if not school:
        return RedirectResponse("/okul-portal", status_code=303)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM inspection_metrics WHERE school_id = ?", (school["id"],))
    metrics = cursor.fetchone()
    conn.close()

    completion = evaluate_school_data(dict(school), dict(metrics) if metrics else {})
    now_str = datetime.now().strftime("%d.%m.%Y")

    return templates.TemplateResponse(request=request, name="school_portal_print.html", context={
        "school": school,
        "metrics": metrics,
        "completion": completion,
        "now_str": now_str
    })

@app.post("/okul-portal/kaydet")
async def school_portal_save(request: Request):
    school = get_current_school(request)
    if not school:
        return RedirectResponse("/okul-portal", status_code=303)

    form_data = await request.form()
    conn = get_db()
    cursor = conn.cursor()

    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
    mudur_ad_soyad = form_data.get("mudur_ad_soyad", "").strip()
    mudur_tel = form_data.get("mudur_tel", "").strip()
    isveren_vekili_ad_soyad = form_data.get("isveren_vekili_ad_soyad", "").strip()
    isveren_vekili_tel = form_data.get("isveren_vekili_tel", "").strip()
    ogrenci_sayisi = int(form_data.get("ogrenci_sayisi", 0) or 0)
    ogretmen_sayisi = int(form_data.get("ogretmen_sayisi", 0) or 0)
    derslik_sayisi = int(form_data.get("derslik_sayisi", 0) or 0)
    bina_sayisi = int(form_data.get("bina_sayisi", 1) or 1)
    eposta = form_data.get("eposta", "").strip()
    telefon = form_data.get("telefon", "").strip()

    # 1. schools tablosunu güncelle
    cursor.execute("""
        UPDATE schools
        SET mudur_ad_soyad = ?,
            mudur_tel = ?,
            isveren_vekili_ad_soyad = ?,
            isveren_vekili_tel = ?,
            ogrenci_sayisi = ?,
            ogretmen_sayisi = ?,
            derslik_sayisi = ?,
            bina_sayisi = ?,
            eposta = ?,
            telefon = ?,
            son_veri_giris_tarihi = ?,
            son_veri_giren_kisi = ?
        WHERE id = ?
    """, (
        mudur_ad_soyad, mudur_tel, isveren_vekili_ad_soyad, isveren_vekili_tel,
        ogrenci_sayisi, ogretmen_sayisi, derslik_sayisi, bina_sayisi, eposta, telefon,
        now_str, mudur_ad_soyad, school["id"]
    ))

    kazan_durumu = form_data.get("kazan_durumu", "Var")
    asansor_durumu = form_data.get("asansor_durumu", "Yok")
    su_deposu_durumu = form_data.get("su_deposu_durumu", "Var")
    paratoner_durumu = form_data.get("paratoner_durumu", "Var")

    cursor.execute("""
        UPDATE inspection_metrics
        SET kamera_sistemi_var_mi = ?,
            toplam_kamera_sayisi = ?,
            kor_nokta_kamera_eksigi = ?,
            kgys_entegre_mi = ?,
            kamera_30gun_kayit_ve_aci_tam_mi = ?,
            ozel_guvenlik_sayisi = ?,
            bahce_gorevlisi_sayisi = ?,
            giris_cikis_kapi_durumu = ?,
            turnike_var_mi = ?,
            ziyaretci_kayit_sistemi = ?,
            ozel_arac_bahce_engeli = ?,
            adep_var_mi = ?,
            adep_tarihi = ?,
            ek7_olcek_var_mi = ?,
            ek7_tarihi = ?,
            tatbikat_turu = ?,
            tatbikat_tarihi = ?,
            tatbikat_katilimci_sayisi = ?,
            guvenlik_ekibi_toplanti_sayisi = ?,
            kazan_durumu = ?,
            asansor_durumu = ?,
            su_deposu_durumu = ?,
            paratoner_durumu = ?,
            kazan_kontrol_tarihi = ?,
            elektrik_topraklama_tarihi = ?,
            paratoner_kontrol_tarihi = ?,
            yangin_tupleri_kontrol_tarihi = ?,
            asansor_muayene_tarihi = ?,
            su_deposu_temizlik_tarihi = ?,
            isg_kurulu_var_mi = ?,
            calisan_temsilcisi_var_mi = ?,
            ilkyardimci_personel_sayisi = ?,
            egitim_yonetici_alan = ?,
            egitim_ogretmen_alan = ?,
            egitim_ogrenci_alan = ?,
            egitim_diger_alan = ?,
            guvenlik_olay_sayisi = ?,
            ihmal_istismar_sayisi = ?,
            bildirilen_ihmal_istismar_sayisi = ?,
            updated_at = CURRENT_TIMESTAMP
        WHERE school_id = ?
    """, (
        form_data.get("kamera_sistemi_var_mi", "Var"),
        int(form_data.get("toplam_kamera_sayisi", 16) or 0),
        int(form_data.get("kor_nokta_kamera_eksigi", 0) or 0),
        form_data.get("kgys_entegre_mi", "Hayır"),
        form_data.get("kamera_30gun_kayit_ve_aci_tam_mi", "Evet"),
        int(form_data.get("ozel_guvenlik_sayisi", 0) or 0),
        int(form_data.get("bahce_gorevlisi_sayisi", 0) or 0),
        form_data.get("giris_cikis_kapi_durumu", "Tek Kapıdan Girişi Olan Okul"),
        form_data.get("turnike_var_mi", "Yok"),
        form_data.get("ziyaretci_kayit_sistemi", "Evet"),
        form_data.get("ozel_arac_bahce_engeli", "Evet"),
        form_data.get("adep_var_mi", "Var"),
        form_data.get("adep_tarihi", ""),
        form_data.get("ek7_olcek_var_mi", "Var"),
        form_data.get("ek7_tarihi", ""),
        form_data.get("tatbikat_turu", "Yangın ve Tahliye"),
        form_data.get("tatbikat_tarihi", ""),
        int(form_data.get("tatbikat_katilimci_sayisi", 0) or 0),
        int(form_data.get("guvenlik_ekibi_toplanti_sayisi", 1) or 0),
        kazan_durumu,
        asansor_durumu,
        su_deposu_durumu,
        paratoner_durumu,
        form_data.get("kazan_kontrol_tarihi", "") if kazan_durumu == "Var" else "",
        form_data.get("elektrik_topraklama_tarihi", ""),
        form_data.get("paratoner_kontrol_tarihi", "") if paratoner_durumu == "Var" else "",
        form_data.get("yangin_tupleri_kontrol_tarihi", ""),
        form_data.get("asansor_muayene_tarihi", "") if asansor_durumu == "Var" else "",
        form_data.get("su_deposu_temizlik_tarihi", "") if su_deposu_durumu == "Var" else "",
        form_data.get("isg_kurulu_var_mi", "Evet"),
        form_data.get("calisan_temsilcisi_var_mi", "Evet"),
        int(form_data.get("ilkyardimci_personel_sayisi", 2) or 0),
        int(form_data.get("egitim_yonetici_alan", 0) or 0),
        int(form_data.get("egitim_ogretmen_alan", 0) or 0),
        int(form_data.get("egitim_ogrenci_alan", 0) or 0),
        int(form_data.get("egitim_diger_alan", 0) or 0),
        int(form_data.get("guvenlik_olay_sayisi", 0) or 0),
        int(form_data.get("ihmal_istismar_sayisi", 0) or 0),
        int(form_data.get("bildirilen_ihmal_istismar_sayisi", 0) or 0),
        school["id"]
    ))

    # 3. Periyodik Kontroller tablosunu senkronize et
    controls_map = [
        ("Kalorifer Kazanı", form_data.get("kazan_kontrol_tarihi", ""), kazan_durumu),
        ("Elektrik İç Tesisat ve Topraklama", form_data.get("elektrik_topraklama_tarihi", ""), "Var"),
        ("Paratoner (Yıldırımlık)", form_data.get("paratoner_kontrol_tarihi", ""), paratoner_durumu),
        ("Yangın Söndürme Tüpleri (YSC)", form_data.get("yangin_tupleri_kontrol_tarihi", ""), "Var"),
        ("Asansör Periyodik Kontrolü", form_data.get("asansor_muayene_tarihi", ""), asansor_durumu)
    ]
    for ekipman, son_tarih, durum in controls_map:
        if durum == "Yok":
            cursor.execute("SELECT id FROM periodic_controls WHERE school_id = ? AND ekipman_turu = ?", (school["id"], ekipman))
            existing_pc = cursor.fetchone()
            if existing_pc:
                cursor.execute("""
                    UPDATE periodic_controls
                    SET son_kontrol_tarihi = 'Muaf / Ekipman Yok',
                        alarm_seviyesi = 'Gecerli'
                    WHERE id = ?
                """, (existing_pc["id"],))
        elif son_tarih:
            cursor.execute("SELECT id FROM periodic_controls WHERE school_id = ? AND ekipman_turu = ?", (school["id"], ekipman))
            existing_pc = cursor.fetchone()
            if existing_pc:
                cursor.execute("""
                    UPDATE periodic_controls
                    SET son_kontrol_tarihi = ?,
                        alarm_seviyesi = 'Gecerli'
                    WHERE id = ?
                """, (son_tarih, existing_pc["id"]))
            else:
                cursor.execute("""
                    INSERT INTO periodic_controls (school_id, ekipman_turu, son_kontrol_tarihi, gecerlilik_ayi, sonraki_kontrol_tarihi, alarm_seviyesi)
                    VALUES (?, ?, ?, 12, ?, 'Gecerli')
                """, (school["id"], ekipman, son_tarih, son_tarih))

    conn.commit()
    conn.close()

    return RedirectResponse("/okul-portal/form?saved=1", status_code=303)

@app.get("/okul-portal/cikis")
def school_portal_logout():
    res = RedirectResponse("/okul-portal", status_code=303)
    res.delete_cookie("session_school_id")
    return res

# ----------------- İLÇE MEM: OKUL VERİ GİRİŞ LİNKLERİ YÖNETİMİ -----------------

@app.get("/okul-baglantilari", response_class=HTMLResponse)
def school_links_admin(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, kurum_kodu, kurum_adi, kurum_statusu, kurum_turu,
               mudur_ad_soyad, mudur_tel, portal_sifre, access_token, son_veri_giris_tarihi
        FROM schools
        ORDER BY oncelik_derecesi ASC, kurum_adi ASC
    """)
    schools = cursor.fetchall()

    cursor.execute("SELECT COUNT(*) FROM schools")
    total_count = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM schools WHERE son_veri_giris_tarihi IS NOT NULL AND son_veri_giris_tarihi != ''")
    updated_count = cursor.fetchone()[0]

    conn.close()

    base_url = str(request.base_url).rstrip('/')

    return templates.TemplateResponse(request=request, name="school_links_admin.html", context={
        "user": user,
        "schools": schools,
        "total_count": total_count,
        "updated_count": updated_count,
        "base_url": base_url,
        "active_page": "links"
    })

@app.get("/export/okul-linkleri")
def export_school_links(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT kurum_kodu, kurum_adi, kurum_statusu, kurum_turu,
               mudur_ad_soyad, mudur_tel, isveren_vekili_ad_soyad, isveren_vekili_tel,
               portal_sifre, access_token, son_veri_giris_tarihi
        FROM schools
        ORDER BY oncelik_derecesi ASC, kurum_adi ASC
    """)
    rows = cursor.fetchall()
    conn.close()

    base_url = str(request.base_url).rstrip('/')

    output = io.StringIO()
    output.write('\ufeff')
    writer = csv.writer(output, delimiter=';')

    writer.writerow([
        "Kurum Kodu", "Kurum Adı", "Statü", "Tür",
        "Okul Müdürü", "Müdür Cep Tel", "İSG Md. Yrd.", "İSG Md. Yrd. Tel",
        "Kurum Şifresi", "Doğrudan Veri Giriş Linki (Şifresiz)", "Son Veri Giriş Durumu"
    ])

    for r in rows:
        direct_link = f"{base_url}/okul-portal/giris/{r['access_token']}"
        writer.writerow([
            r["kurum_kodu"], r["kurum_adi"], r["kurum_statusu"], r["kurum_turu"],
            r["mudur_ad_soyad"] or "", r["mudur_tel"] or "",
            r["isveren_vekili_ad_soyad"] or "", r["isveren_vekili_tel"] or "",
            r["portal_sifre"] or "isg2026",
            direct_link,
            r["son_veri_giris_tarihi"] or "Henüz Girilmedi"
        ])

    output.seek(0)
    filename = f"Muratpasa_Okul_Veri_Giris_Linkleri_{datetime.now().strftime('%Y%m%d')}.csv"
    return StreamingResponse(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# ----------------- İGU GÖREVLENDİRME & ÇALIŞMA DAĞILIM ÇİZELGESİ -----------------

@app.get("/igu-gorevlendirme", response_class=HTMLResponse)
def igu_assignments_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, username, full_name, role, title, email, tc_kimlik, kadro_yeri, igu_sinifi, gorev_turu
        FROM users
        ORDER BY 
            CASE WHEN role = 'koordinator' THEN 0 ELSE 1 END,
            igu_sinifi ASC,
            full_name ASC
    """)
    experts = cursor.fetchall()

    experts_data = []
    total_assigned_schools = 0
    total_personnel = 0
    total_minutes = 0

    for exp in experts:
        cursor.execute("""
            SELECT id, kurum_kodu, kurum_adi, kurum_statusu, kurum_turu, tehlike_sinifi,
                   personel_sayisi, aylik_hizmet_suresi_dk, mudur_ad_soyad, mudur_tel
            FROM schools
            WHERE sorumlu_uzman_id = ?
            ORDER BY oncelik_derecesi ASC, kurum_adi ASC
        """, (exp["id"],))
        assigned_schools = cursor.fetchall()

        exp_personnel = sum(s["personel_sayisi"] or 0 for s in assigned_schools)
        exp_minutes = sum(s["aylik_hizmet_suresi_dk"] or 0 for s in assigned_schools)

        total_assigned_schools += len(assigned_schools)
        total_personnel += exp_personnel
        total_minutes += exp_minutes

        experts_data.append({
            "id": exp["id"],
            "username": exp["username"],
            "full_name": exp["full_name"],
            "role": exp["role"],
            "title": exp["title"],
            "email": exp["email"],
            "tc_kimlik": exp["tc_kimlik"],
            "kadro_yeri": exp["kadro_yeri"],
            "igu_sinifi": exp["igu_sinifi"] or "İGU",
            "gorev_turu": exp["gorev_turu"] or "Tam Süreli",
            "schools": assigned_schools,
            "total_personnel": exp_personnel,
            "total_minutes": exp_minutes
        })

    # Atanmamış Kurumlar (Genel Havuz)
    cursor.execute("""
        SELECT id, kurum_kodu, kurum_adi, kurum_statusu, kurum_turu, tehlike_sinifi, oncelik_derecesi,
               personel_sayisi, mudur_ad_soyad, mudur_tel, isveren_vekili_ad_soyad, isveren_vekili_tel
        FROM schools
        WHERE sorumlu_uzman_id IS NULL
        ORDER BY oncelik_derecesi ASC, kurum_adi ASC
    """)
    unassigned_all = cursor.fetchall()
    unassigned_official = [s for s in unassigned_all if s["kurum_statusu"] == 'Resmî']
    unassigned_private = [s for s in unassigned_all if s["kurum_statusu"] == 'Özel']

    conn.close()

    return templates.TemplateResponse(request=request, name="igu_assignments.html", context={
        "user": user,
        "experts_data": experts_data,
        "total_experts": len(experts),
        "total_assigned_schools": total_assigned_schools,
        "total_personnel": total_personnel,
        "total_minutes": total_minutes,
        "unassigned_official": unassigned_official,
        "unassigned_private": unassigned_private,
        "total_unassigned": len(unassigned_all),
        "active_page": "igu"
    })

@app.get("/export/igu-gorevlendirme")
def export_igu_assignments(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT u.full_name as uzman_adi, u.igu_sinifi, u.kadro_yeri, u.tc_kimlik as uzman_tc,
               s.kurum_kodu, s.kurum_adi, s.kurum_statusu, s.kurum_turu, s.tehlike_sinifi,
               s.personel_sayisi, s.aylik_hizmet_suresi_dk,
               s.mudur_ad_soyad, s.mudur_tel
        FROM schools s
        JOIN users u ON s.sorumlu_uzman_id = u.id
        ORDER BY u.id ASC, s.kurum_adi ASC
    """)
    rows = cursor.fetchall()
    conn.close()

    output = io.StringIO()
    output.write('\ufeff')
    writer = csv.writer(output, delimiter=';')

    writer.writerow([
        "İSG Uzmanı Adı Soyadı", "İGU Sınıfı", "Kadro / Görev Yeri", "Uzman TC Kimlik",
        "Kurum Kodu", "Kurum Adı", "Statü", "Kurum Türü", "Tehlike Sınıfı",
        "Personel Sayısı", "Aylık Yasal Hizmet Süresi (Dk)",
        "Okul Müdürü", "Müdür Telefon"
    ])

    for r in rows:
        writer.writerow(list(r))

    output.seek(0)
    filename = f"Muratpasa_IGU_Gorevlendirme_Cizelgesi_{datetime.now().strftime('%Y%m%d')}.csv"
    return StreamingResponse(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# ----------------- DİNAMİK İSG RAPORLAMA & VERİ FİLTRELEME SİHİRBAZI -----------------

def build_report_query(filters: dict):
    conditions = []
    params = []
    labels = []

    # Turnike
    if filters.get("turnike") and filters["turnike"] != "all":
        conditions.append("im.turnike_var_mi = ?")
        params.append(filters["turnike"])
        labels.append(f"Turnike: {filters['turnike']}")

    # El/Kapı Dedektörü
    if filters.get("el_dedektoru") and filters["el_dedektoru"] != "all":
        conditions.append("(im.el_dedektoru_var_mi = ? OR im.kapi_dedektoru_var_mi = ?)")
        params.extend([filters["el_dedektoru"], filters["el_dedektoru"]])
        labels.append(f"Dedektör: {filters['el_dedektoru']}")

    # Özel Güvenlik
    if filters.get("ozel_guvenlik") == "var":
        conditions.append("im.ozel_guvenlik_sayisi > 0")
        labels.append("Özel Güvenlik: Var (≥1)")
    elif filters.get("ozel_guvenlik") == "yok":
        conditions.append("(im.ozel_guvenlik_sayisi IS NULL OR im.ozel_guvenlik_sayisi = 0)")
        labels.append("Özel Güvenlik: Yok")

    # Bahçe Görevlisi (TYP)
    if filters.get("bahce_gorevlisi") == "var":
        conditions.append("im.bahce_gorevlisi_sayisi > 0")
        labels.append("Bahçe Görevlisi: Var")
    elif filters.get("bahce_gorevlisi") == "yok":
        conditions.append("(im.bahce_gorevlisi_sayisi IS NULL OR im.bahce_gorevlisi_sayisi = 0)")
        labels.append("Bahçe Görevlisi: Yok")

    # Sertifikalı Bahçe Görevlisi
    if filters.get("sertifikali_bahce") == "var":
        conditions.append("im.sertifikali_bahce_gorevlisi_sayisi > 0")
        labels.append("Sertifikalı Bahçe G.: Var")
    elif filters.get("sertifikali_bahce") == "yok":
        conditions.append("(im.sertifikali_bahce_gorevlisi_sayisi IS NULL OR im.sertifikali_bahce_gorevlisi_sayisi = 0)")
        labels.append("Sertifikalı Bahçe G.: Yok")

    # Ziyaretçi Kayıt
    if filters.get("ziyaretci_kayit") and filters["ziyaretci_kayit"] != "all":
        conditions.append("im.ziyaretci_kayit_sistemi = ?")
        params.append(filters["ziyaretci_kayit"])
        labels.append(f"Ziyaretçi Defteri: {filters['ziyaretci_kayit']}")

    # Kamera
    if filters.get("kamera_durumu") and filters["kamera_durumu"] != "all":
        conditions.append("im.kamera_sistemi_var_mi = ?")
        params.append(filters["kamera_durumu"])
        labels.append(f"Kamera: {filters['kamera_durumu']}")

    # KGYS
    if filters.get("kgys") and filters["kgys"] != "all":
        conditions.append("im.kgys_entegre_mi = ?")
        params.append(filters["kgys"])
        labels.append(f"KGYS: {filters['kgys']}")

    # Kapı Durumu
    if filters.get("kapi_durumu") and filters["kapi_durumu"] != "all":
        conditions.append("im.giris_cikis_kapi_durumu = ?")
        params.append(filters["kapi_durumu"])
        labels.append(f"Kapı: {filters['kapi_durumu']}")

    # Bahçeye Araç Engeli
    if filters.get("arac_engeli") and filters["arac_engeli"] != "all":
        conditions.append("im.ozel_arac_bahce_engeli = ?")
        params.append(filters["arac_engeli"])
        labels.append(f"Bahçeye Araç Engeli: {filters['arac_engeli']}")

    # Kurum Statüsü
    if filters.get("kurum_statusu") and filters["kurum_statusu"] != "all":
        conditions.append("s.kurum_statusu = ?")
        params.append(filters["kurum_statusu"])
        labels.append(f"Statü: {filters['kurum_statusu']}")

    # Kurum Türü
    if filters.get("kurum_turu") and filters["kurum_turu"] != "all":
        conditions.append("s.kurum_turu LIKE ?")
        params.append(f"%{filters['kurum_turu']}%")
        labels.append(f"Tür: {filters['kurum_turu']}")

    # Öncelik Derecesi
    if filters.get("oncelik_derecesi") and filters["oncelik_derecesi"] != "all":
        conditions.append("s.oncelik_derecesi = ?")
        params.append(int(filters["oncelik_derecesi"]))
        labels.append(f"Öncelik: {filters['oncelik_derecesi']}. Derece")

    # Sorumlu Uzman
    if filters.get("sorumlu_uzman_id") and filters["sorumlu_uzman_id"] != "all":
        if filters["sorumlu_uzman_id"] == "atanmadi":
            conditions.append("s.sorumlu_uzman_id IS NULL")
            labels.append("Uzman: Atanmadı (Genel Havuz)")
        else:
            conditions.append("s.sorumlu_uzman_id = ?")
            params.append(int(filters["sorumlu_uzman_id"]))
            labels.append(f"Uzman ID: {filters['sorumlu_uzman_id']}")

    # ADEP
    if filters.get("adep") and filters["adep"] != "all":
        conditions.append("im.adep_var_mi = ?")
        params.append(filters["adep"])
        labels.append(f"ADEP: {filters['adep']}")

    # Tatbikat
    if filters.get("tatbikat") == "yapildi":
        conditions.append("im.tatbikat_tarihi IS NOT NULL AND im.tatbikat_tarihi != ''")
        labels.append("Tatbikat: Yapıldı")
    elif filters.get("tatbikat") == "yapilmadi":
        conditions.append("(im.tatbikat_tarihi IS NULL OR im.tatbikat_tarihi = '')")
        labels.append("Tatbikat: Eksik / Yapılmadı")

    # Kazan
    if filters.get("kazan") == "yapildi":
        conditions.append("im.kazan_kontrol_tarihi IS NOT NULL AND im.kazan_kontrol_tarihi != ''")
        labels.append("Kazan: Muayeneli")
    elif filters.get("kazan") == "eksik":
        conditions.append("(im.kazan_kontrol_tarihi IS NULL OR im.kazan_kontrol_tarihi = '')")
        labels.append("Kazan: Muayene Eksik")

    # Elektrik
    if filters.get("elektrik") == "yapildi":
        conditions.append("im.elektrik_topraklama_tarihi IS NOT NULL AND im.elektrik_topraklama_tarihi != ''")
        labels.append("Elektrik: Ölçüldü")
    elif filters.get("elektrik") == "eksik":
        conditions.append("(im.elektrik_topraklama_tarihi IS NULL OR im.elektrik_topraklama_tarihi = '')")
        labels.append("Elektrik: Ölçüm Eksik")

    where_sql = ("WHERE " + " AND ".join(conditions)) if conditions else ""
    return where_sql, params, labels

@app.get("/dinamik-rapor", response_class=HTMLResponse)
def dynamic_report_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    query_params = dict(request.query_params)
    where_sql, params, active_filter_labels = build_report_query(query_params)

    conn = get_db()
    cursor = conn.cursor()

    # Toplam kurum sayısı
    cursor.execute("SELECT COUNT(*) FROM schools")
    total_schools = cursor.fetchone()[0]

    # Uzman listesi (filtre için)
    cursor.execute("SELECT id, full_name, igu_sinifi FROM users ORDER BY full_name ASC")
    experts = cursor.fetchall()

    # Filtrelenen okulları çek
    sql = f"""
    SELECT s.id, s.kurum_kodu, s.kurum_adi, s.kurum_statusu, s.kurum_turu,
           s.oncelik_derecesi, s.tehlike_sinifi, s.personel_sayisi,
           s.mudur_ad_soyad, s.mudur_tel, s.isveren_vekili_ad_soyad, s.isveren_vekili_tel,
           u.full_name as uzman_adi,
           im.turnike_var_mi, im.el_dedektoru_var_mi, im.kapi_dedektoru_var_mi,
           im.ozel_guvenlik_sayisi, im.bahce_gorevlisi_sayisi, im.sertifikali_bahce_gorevlisi_sayisi,
           im.kamera_sistemi_var_mi, im.toplam_kamera_sayisi, im.kgys_entegre_mi,
           im.giris_cikis_kapi_durumu, im.ziyaretci_kayit_sistemi, im.ozel_arac_bahce_engeli,
           im.adep_var_mi, im.adep_tarihi, im.ek7_olcek_var_mi, im.ek7_tarihi,
           im.tatbikat_tarihi, im.kazan_kontrol_tarihi, im.elektrik_topraklama_tarihi,
           im.yangin_tupleri_kontrol_tarihi, im.asansor_muayene_tarihi
    FROM schools s
    LEFT JOIN users u ON s.sorumlu_uzman_id = u.id
    LEFT JOIN inspection_metrics im ON s.id = im.school_id
    {where_sql}
    ORDER BY s.oncelik_derecesi ASC, s.kurum_adi ASC
    """
    cursor.execute(sql, params)
    filtered_schools = cursor.fetchall()
    conn.close()

    raw_query = str(request.query_params)

    return templates.TemplateResponse(request=request, name="dynamic_report.html", context={
        "user": user,
        "filtered_schools": filtered_schools,
        "total_schools": total_schools,
        "experts": experts,
        "current_filters": query_params,
        "active_filter_labels": active_filter_labels,
        "query_string": raw_query,
        "active_page": "reports"
    })

@app.get("/export/dinamik-rapor")
def export_dynamic_report(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    query_params = dict(request.query_params)
    where_sql, params, active_filter_labels = build_report_query(query_params)

    conn = get_db()
    cursor = conn.cursor()

    sql = f"""
    SELECT s.kurum_kodu, s.kurum_adi, s.kurum_statusu, s.kurum_turu,
           s.oncelik_derecesi, s.tehlike_sinifi, s.personel_sayisi,
           s.mudur_ad_soyad, s.mudur_tel, s.isveren_vekili_ad_soyad, s.isveren_vekili_tel,
           u.full_name as uzman_adi,
           im.turnike_var_mi, im.el_dedektoru_var_mi,
           im.ozel_guvenlik_sayisi, im.bahce_gorevlisi_sayisi, im.sertifikali_bahce_gorevlisi_sayisi,
           im.kamera_sistemi_var_mi, im.toplam_kamera_sayisi, im.kgys_entegre_mi,
           im.giris_cikis_kapi_durumu, im.ziyaretci_kayit_sistemi, im.ozel_arac_bahce_engeli,
           im.adep_var_mi, im.adep_tarihi, im.tatbikat_tarihi,
           im.kazan_kontrol_tarihi, im.elektrik_topraklama_tarihi, im.yangin_tupleri_kontrol_tarihi
    FROM schools s
    LEFT JOIN users u ON s.sorumlu_uzman_id = u.id
    LEFT JOIN inspection_metrics im ON s.id = im.school_id
    {where_sql}
    ORDER BY s.oncelik_derecesi ASC, s.kurum_adi ASC
    """
    cursor.execute(sql, params)
    rows = cursor.fetchall()
    conn.close()

    output = io.StringIO()
    output.write('\ufeff')
    writer = csv.writer(output, delimiter=';')

    writer.writerow([
        "Kurum Kodu", "Kurum Adı", "Statü", "Kurum Türü",
        "Öncelik Derecesi", "Tehlike Sınıfı", "Personel Sayısı",
        "Okul Müdürü", "Müdür Telefon", "İSG Md. Yrd.", "İSG Md. Yrd. Telefon",
        "Sorumlu İGU Uzmanı", "Turnike Durumu", "Dedektör Durumu",
        "Özel Güvenlik Sayısı", "Bahçe Görevlisi (TYP)", "Sertifikalı Bahçe Görevlisi",
        "Kamera Sistemi", "Toplam Kamera", "KGYS Entegrasyonu",
        "Giriş Kapısı", "Ziyaretçi Defteri", "Bahçeye Araç Engeli",
        "ADEP Durumu", "ADEP Tarihi", "Son Tatbikat Tarihi",
        "Kazan Kontrol Tarihi", "Elektrik Topraklama Tarihi", "Yangın Tüpü Kontrol Tarihi"
    ])

    for r in rows:
        writer.writerow(list(r))

    output.seek(0)
    filename = f"Muratpasa_ISG_Ozel_Rapor_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
    return StreamingResponse(
        io.BytesIO(output.getvalue().encode('utf-8-sig')),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/rapor/yazdir", response_class=HTMLResponse)
def print_report_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    query_params = dict(request.query_params)
    where_sql, params, active_filter_labels = build_report_query(query_params)

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM schools")
    total_schools = cursor.fetchone()[0]

    sql = f"""
    SELECT s.id, s.kurum_kodu, s.kurum_adi, s.kurum_statusu, s.kurum_turu,
           s.oncelik_derecesi, s.tehlike_sinifi, s.personel_sayisi,
           s.mudur_ad_soyad, s.mudur_tel, s.isveren_vekili_ad_soyad, s.isveren_vekili_tel,
           u.full_name as uzman_adi,
           im.turnike_var_mi, im.el_dedektoru_var_mi, im.kapi_dedektoru_var_mi,
           im.ozel_guvenlik_sayisi, im.bahce_gorevlisi_sayisi,
           im.kamera_sistemi_var_mi, im.toplam_kamera_sayisi, im.kgys_entegre_mi
    FROM schools s
    LEFT JOIN users u ON s.sorumlu_uzman_id = u.id
    LEFT JOIN inspection_metrics im ON s.id = im.school_id
    {where_sql}
    ORDER BY s.oncelik_derecesi ASC, s.kurum_adi ASC
    """
    cursor.execute(sql, params)
    filtered_schools = cursor.fetchall()
    conn.close()

    now_str = datetime.now().strftime("%d.%m.%Y")

    return templates.TemplateResponse(request=request, name="report_print.html", context={
        "user": user,
        "filtered_schools": filtered_schools,
        "total_schools": total_schools,
        "active_filter_labels": active_filter_labels,
        "now_str": now_str
    })

# ----------------- OKUL İSG VERİ TAMAMLAMA & EKSİKLİK ANALİZİ -----------------

@app.get("/veri-tamamlama", response_class=HTMLResponse)
def data_completion_dashboard(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()

    # Uzmanlar listesi
    cursor.execute("SELECT id, full_name, igu_sinifi FROM users WHERE role = 'uzman' ORDER BY full_name ASC")
    experts = cursor.fetchall()

    # Tüm okulları ve metriklerini çek
    cursor.execute("""
    SELECT s.*, 
           u.full_name as uzman_adi,
           im.adep_var_mi, im.adep_tarihi, im.ek7_olcek_var_mi, im.ek7_tarihi,
           im.kamera_sistemi_var_mi, im.toplam_kamera_sayisi, im.kamera_30gun_kayit_ve_aci_tam_mi,
           im.ziyaretci_kayit_sistemi,
           im.tatbikat_tarihi, im.kazan_kontrol_tarihi, im.elektrik_topraklama_tarihi,
           im.paratoner_kontrol_tarihi, im.yangin_tupleri_kontrol_tarihi,
           im.calisan_temsilcisi_var_mi, im.ilkyardimci_personel_sayisi,
           im.sifir_atik_belge_no, im.guvenlik_izin_belgesi_var_mi
    FROM schools s
    LEFT JOIN users u ON s.sorumlu_uzman_id = u.id
    LEFT JOIN inspection_metrics im ON s.id = im.school_id
    ORDER BY s.oncelik_derecesi ASC, s.kurum_adi ASC
    """)
    rows = cursor.fetchall()
    conn.close()

    # Query parametreleri
    q = request.query_params.get("q", "").strip().lower()
    r_filter = request.query_params.get("range", "all")
    missing_field = request.query_params.get("missing_field", "all")
    expert_id = request.query_params.get("expert_id", "all")
    status = request.query_params.get("status", "all")
    school_type = request.query_params.get("school_type", "all")

    # İlçe Genel Metrikleri
    total_schools = len(rows)
    all_evaluations = []
    cat_district_totals = {}

    for row in rows:
        eval_res = evaluate_school_data(row, row)
        all_evaluations.append((row, eval_res))
        for cat, c_data in eval_res["categories"].items():
            if cat not in cat_district_totals:
                cat_district_totals[cat] = {"completed": 0, "total": 0}
            cat_district_totals[cat]["completed"] += c_data["completed"]
            cat_district_totals[cat]["total"] += c_data["total"]

    district_avg_pct = round(sum(e[1]["percentage"] for e in all_evaluations) / total_schools) if total_schools > 0 else 0
    high_count = sum(1 for e in all_evaluations if e[1]["percentage"] >= 80)
    medium_count = sum(1 for e in all_evaluations if 50 <= e[1]["percentage"] < 80)
    critical_count = sum(1 for e in all_evaluations if e[1]["percentage"] < 50)
    complete_count = sum(1 for e in all_evaluations if e[1]["percentage"] == 100)

    for cat in cat_district_totals:
        tot = cat_district_totals[cat]["total"]
        comp = cat_district_totals[cat]["completed"]
        cat_district_totals[cat]["pct"] = round((comp / tot) * 100) if tot > 0 else 0

    # Filtreleme
    filtered_list = []
    for row, eval_res in all_evaluations:
        if q and (q not in row["kurum_adi"].lower() and q not in str(row["kurum_kodu"]).lower()):
            continue

        if r_filter == "critical" and eval_res["percentage"] >= 50:
            continue
        elif r_filter == "medium" and not (50 <= eval_res["percentage"] < 80):
            continue
        elif r_filter == "high" and eval_res["percentage"] < 80:
            continue
        elif r_filter == "complete" and eval_res["percentage"] != 100:
            continue
        elif r_filter == "missing" and eval_res["percentage"] == 100:
            continue

        if missing_field != "all":
            missing_ids = [item["id"] for item in eval_res["missing_items"]]
            if missing_field not in missing_ids:
                continue

        if expert_id != "all":
            if expert_id == "atanmadi":
                if row["sorumlu_uzman_id"] is not None:
                    continue
            elif str(row["sorumlu_uzman_id"]) != expert_id:
                continue

        if status != "all" and row["kurum_statusu"] != status:
            continue

        if school_type != "all" and school_type not in (row["kurum_turu"] or ""):
            continue

        filtered_list.append({
            "school": row,
            "completion": eval_res
        })

    return templates.TemplateResponse(request=request, name="data_completion.html", context={
        "user": user,
        "schools": filtered_list,
        "total_schools": total_schools,
        "district_avg_pct": district_avg_pct,
        "high_count": high_count,
        "medium_count": medium_count,
        "critical_count": critical_count,
        "complete_count": complete_count,
        "cat_district_totals": cat_district_totals,
        "experts": experts,
        "criteria_list": EVALUATION_CRITERIA,
        "filters": {
            "q": q,
            "range": r_filter,
            "missing_field": missing_field,
            "expert_id": expert_id,
            "status": status,
            "school_type": school_type
        },
        "query_string": str(request.query_params),
        "active_page": "completion"
    })

@app.get("/export/veri-tamamlama")
def export_data_completion(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT s.*, 
           u.full_name as uzman_adi,
           im.adep_var_mi, im.adep_tarihi, im.ek7_olcek_var_mi, im.ek7_tarihi,
           im.kamera_sistemi_var_mi, im.toplam_kamera_sayisi, im.kamera_30gun_kayit_ve_aci_tam_mi,
           im.ziyaretci_kayit_sistemi,
           im.tatbikat_tarihi, im.kazan_kontrol_tarihi, im.elektrik_topraklama_tarihi,
           im.paratoner_kontrol_tarihi, im.yangin_tupleri_kontrol_tarihi,
           im.calisan_temsilcisi_var_mi, im.ilkyardimci_personel_sayisi,
           im.sifir_atik_belge_no, im.guvenlik_izin_belgesi_var_mi
    FROM schools s
    LEFT JOIN users u ON s.sorumlu_uzman_id = u.id
    LEFT JOIN inspection_metrics im ON s.id = im.school_id
    ORDER BY s.oncelik_derecesi ASC, s.kurum_adi ASC
    """)
    rows = cursor.fetchall()
    conn.close()

    q = request.query_params.get("q", "").strip().lower()
    r_filter = request.query_params.get("range", "all")
    missing_field = request.query_params.get("missing_field", "all")
    expert_id = request.query_params.get("expert_id", "all")
    status = request.query_params.get("status", "all")
    school_type = request.query_params.get("school_type", "all")

    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")

    writer.writerow([
        "Kurum Kodu",
        "Kurum Adı",
        "Statü",
        "Kurum Türü",
        "Öncelik Derecesi",
        "Sorumlu İGU Uzmanı",
        "Müdür Adı Soyadı",
        "Müdür Telefon",
        "İşveren Vekili Adı Soyadı",
        "İşveren Vekili Telefon",
        "Veri Tamamlanma Oranı (%)",
        "Tamamlanan Veri Sayısı",
        "Eksik Veri Sayısı",
        "Eksik Veriler",
        "Son Veri Giriş Tarihi"
    ])

    for row in rows:
        eval_res = evaluate_school_data(row, row)

        if q and (q not in row["kurum_adi"].lower() and q not in str(row["kurum_kodu"]).lower()):
            continue
        if r_filter == "critical" and eval_res["percentage"] >= 50:
            continue
        elif r_filter == "medium" and not (50 <= eval_res["percentage"] < 80):
            continue
        elif r_filter == "high" and eval_res["percentage"] < 80:
            continue
        elif r_filter == "complete" and eval_res["percentage"] != 100:
            continue
        elif r_filter == "missing" and eval_res["percentage"] == 100:
            continue

        if missing_field != "all":
            missing_ids = [item["id"] for item in eval_res["missing_items"]]
            if missing_field not in missing_ids:
                continue

        if expert_id != "all":
            if expert_id == "atanmadi":
                if row["sorumlu_uzman_id"] is not None:
                    continue
            elif str(row["sorumlu_uzman_id"]) != expert_id:
                continue

        if status != "all" and row["kurum_statusu"] != status:
            continue
        if school_type != "all" and school_type not in (row["kurum_turu"] or ""):
            continue

        missing_labels = ", ".join([item["label"] for item in eval_res["missing_items"]])

        writer.writerow([
            row["kurum_kodu"],
            row["kurum_adi"],
            row["kurum_statusu"] or "",
            row["kurum_turu"] or "",
            f"{row['oncelik_derecesi']}. Derece" if row["oncelik_derecesi"] else "3. Derece",
            row["uzman_adi"] or "Atanmadı",
            row["mudur_ad_soyad"] or "",
            row["mudur_tel"] or "",
            row["isveren_vekili_ad_soyad"] or "",
            row["isveren_vekili_tel"] or "",
            f"%{eval_res['percentage']}",
            eval_res["completed_count"],
            eval_res["missing_count"],
            missing_labels,
            row["son_veri_giris_tarihi"] or "Giriş Yapılmadı"
        ])

    csv_bytes = "\ufeff" + output.getvalue()
    return Response(
        content=csv_bytes.encode('utf-8-sig'),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename=Muratpasa_Okul_ISG_Veri_Tamamlama_Raporu_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"}
    )

# ----------------- V2.4.0 ENTEGRASYON DİREKTİFİ: 6 TEMEL MODÜL ROUTE'LARI -----------------

# MODÜL 2: ADEP & İSG İÇ YÖNERGESİ ÜRETİM MOTORU
@app.get("/adep-yonerge-motoru", response_class=HTMLResponse)
def adep_engine_page(request: Request, school_id: Optional[int] = None):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, kurum_kodu, kurum_adi, kademe, kurum_turu, tehlike_sinifi, oncelik_derecesi, mudur_ad_soyad, mudur_tel, isveren_vekili_ad_soyad, isveren_vekili_tel, personel_sayisi, ogretmen_sayisi, ogrenci_sayisi FROM schools WHERE is_active_school = 1 ORDER BY oncelik_derecesi ASC, kurum_adi ASC")
    all_schools = c.fetchall()

    selected_school = None
    if school_id:
        c.execute("SELECT * FROM schools WHERE id = ?", (school_id,))
        selected_school = c.fetchone()
    if not selected_school and all_schools:
        c.execute("SELECT * FROM schools WHERE id = ?", (all_schools[0]['id'],))
        selected_school = c.fetchone()

    staff_count = selected_school['personel_sayisi'] if selected_school and selected_school['personel_sayisi'] else 45
    kademe = selected_school['kademe'] if selected_school and selected_school['kademe'] else 'ORTAOKUL'
    quotas = calculate_quotas(kademe, staff_count)
    units = get_architectural_units(kademe)
    conn.close()

    return templates.TemplateResponse(request=request, name="adep_engine.html", context={
        "user": user,
        "all_schools": all_schools,
        "selected_school": selected_school,
        "staff_count": staff_count,
        "quotas": quotas,
        "units": units,
        "active_page": "adep"
    })

@app.get("/api/adep-belge-indir/{school_id}")
def download_adep_docx(school_id: int):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM schools WHERE id = ?", (school_id,))
    school = c.fetchone()
    if not school:
        conn.close()
        raise HTTPException(status_code=404, detail="Okul bulunamadı.")
    c.execute("SELECT * FROM inspection_metrics WHERE school_id = ?", (school_id,))
    metrics = c.fetchone()
    conn.close()

    bio = generate_adep_docx(school, metrics)
    filename = f"2026_ADEP_ve_ISG_Yonergesi_{school['kurum_kodu']}.docx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# MODÜL 3: KRİZ VE AKTİF SALDIRGAN MÜDAHALE PORTALI (LOCKDOWN & ÇDE)
@app.get("/kriz-yonetimi", response_class=HTMLResponse)
def crisis_portal_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT i.*, s.kurum_adi, s.kurum_kodu
        FROM incidents i
        JOIN schools s ON i.school_id = s.id
        ORDER BY i.created_at DESC
    """)
    incidents = c.fetchall()
    c.execute("SELECT count(*) FROM incidents WHERE durum = 'Açık' AND cde_seviye = 'SEVIYE_3_CDE'")
    active_cde = c.fetchone()[0]
    c.execute("SELECT id, kurum_kodu, kurum_adi FROM schools WHERE is_active_school = 1 ORDER BY kurum_adi ASC")
    all_schools = c.fetchall()
    conn.close()

    return templates.TemplateResponse(request=request, name="crisis_portal.html", context={
        "user": user,
        "incidents": incidents,
        "active_cde_count": active_cde,
        "all_schools": all_schools,
        "active_page": "crisis"
    })

@app.post("/kriz-yonetimi/olay-ekle")
def add_incident(school_id: int = Form(...), olay_turu: str = Form(...), cde_seviye: str = Form(...), termin_saat: int = Form(24), bildiren_kisi: str = Form(...), aciklama: str = Form(...), kvkk_onay: int = Form(1)):
    now = datetime.now()
    termin = (now + timedelta(hours=termin_saat)).strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        INSERT INTO incidents (school_id, olay_turu, cde_seviye, bildirim_112_saati, bildiren_kisi, cde_toplanma_termini, aciklama, kvkk_mahremiyet_onayi, durum)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Açık')
    """, (school_id, olay_turu, cde_seviye, now.strftime("%Y-%m-%d %H:%M:%S"), bildiren_kisi, termin, aciklama, kvkk_onay))
    conn.commit()
    conn.close()
    return RedirectResponse("/kriz-yonetimi", status_code=303)

@app.post("/kriz-yonetimi/lockdown-tetikle")
def trigger_lockdown(school_id: int = Form(...)):
    now = datetime.now()
    termin = (now + timedelta(hours=24)).strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        INSERT INTO incidents (school_id, olay_turu, cde_seviye, bildirim_112_saati, bildiren_kisi, cde_toplanma_termini, aciklama, durum)
        VALUES (?, 'SILAHLI_SALDIRI_TEHDIT', 'SEVIYE_3_CDE', ?, 'Sistem Otomatik Lockdown', ?, 'Kilit-Kapıyı Kilitle (Lockdown) protokolü tetiklendi. Acil 112 ve ÇDE devreye alındı.', 'Açık')
    """, (school_id, now.strftime("%Y-%m-%d %H:%M:%S"), termin))
    conn.commit()
    conn.close()
    return RedirectResponse("/kriz-yonetimi", status_code=303)

@app.post("/kriz-yonetimi/olay-durum-guncelle/{incident_id}")
def update_incident_status(incident_id: int, yeni_durum: str = Form(...)):
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE incidents SET durum = ?, olay_sonrasi_degerlendirme_72h = 1 WHERE id = ?", (yeni_durum, incident_id))
    conn.commit()
    conn.close()
    return RedirectResponse("/kriz-yonetimi", status_code=303)

# ==================== YAPAY ZEKA (GEMINI AI) MODÜLLERİ ====================

@app.post("/api/ai/analyze-incident")
async def api_analyze_incident(request: Request):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Oturum açmanız gerekiyor.")

    body = await request.json()
    raw_text = body.get("text", "").strip()
    if not raw_text:
        raise HTTPException(status_code=400, detail="Analiz için metin girilmedi.")

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, kurum_adi, kurum_kodu FROM schools WHERE is_active_school = 1")
    schools = [dict(r) for r in c.fetchall()]
    conn.close()

    try:
        result = isg_ai.analyze_incident_ai(raw_text, schools)
        return {"success": True, "data": result}
    except Exception as e:
        return {"success": False, "error": str(e)}

@app.get("/isg-asistan", response_class=HTMLResponse)
def ai_advisor_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    return templates.TemplateResponse(request=request, name="ai_advisor.html", context={
        "user": user,
        "active_page": "ai_advisor"
    })

@app.post("/api/ai/ask-advisor")
async def api_ask_advisor(request: Request):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Oturum açmanız gerekiyor.")

    body = await request.json()
    question = body.get("question", "").strip()
    if not question:
        raise HTTPException(status_code=400, detail="Soru metni girilmedi.")

    try:
        answer = isg_ai.ask_isg_advisor_ai(question)
        return {"success": True, "answer": answer}
    except Exception as e:
        return {"success": False, "error": str(e)}

@app.get("/ai-veri-aktarimi", response_class=HTMLResponse)
def ai_ingestion_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, kurum_kodu, kurum_adi FROM schools WHERE is_active_school = 1 ORDER BY kurum_adi ASC")
    schools = [dict(r) for r in c.fetchall()]

    c.execute("""
        SELECT d.*, s.kurum_adi, s.kurum_kodu
        FROM documents_archive d
        JOIN schools s ON d.school_id = s.id
        ORDER BY d.id DESC LIMIT 20
    """)
    recent_docs = [dict(r) for r in c.fetchall()]
    conn.close()

    storage_stats = isg_storage.get_storage_statistics()

    return templates.TemplateResponse(request=request, name="ai_ingestion.html", context={
        "user": user,
        "schools": schools,
        "recent_docs": recent_docs,
        "storage_stats": storage_stats,
        "active_page": "ai_ingestion"
    })

@app.post("/api/ai/parse-document")
async def api_parse_document(
    request: Request,
    doc_file: UploadFile = File(...),
    user_instruction: Optional[str] = Form("")
):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Oturum açmanız gerekiyor.")

    file_bytes = await doc_file.read()
    if not file_bytes:
        raise HTTPException(status_code=400, detail="Boş dosya yüklendi.")

    # Dosyayı kaydet
    upload_dir = os.path.join(BASE_DIR, "static", "uploads", "isg_belgeler")
    os.makedirs(upload_dir, exist_ok=True)
    clean_name = os.path.basename(doc_file.filename).replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    saved_filename = f"{timestamp}_{clean_name}"
    file_path = os.path.join(upload_dir, saved_filename)
    with open(file_path, "wb") as f:
        f.write(file_bytes)

    relative_file_path = f"/static/uploads/isg_belgeler/{saved_filename}"

    # Dosya içeriğini çıkar
    file_text, inline_img = isg_ai.extract_file_content(clean_name, file_bytes)

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, kurum_kodu, kurum_adi FROM schools WHERE is_active_school = 1")
    schools = [dict(r) for r in c.fetchall()]
    conn.close()

    try:
        extracted = isg_ai.analyze_and_extract_mebbis_report(file_text, inline_img, user_instruction, schools)
        return {
            "success": True,
            "data": extracted,
            "saved_file": {
                "filename": doc_file.filename,
                "file_path": relative_file_path
            }
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

@app.post("/api/ai/apply-updates")
async def api_apply_updates(request: Request):
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Oturum açmanız gerekiyor.")

    body = await request.json()
    school_id = body.get("school_id")
    if not school_id:
        raise HTTPException(status_code=400, detail="Okul seçilmedi.")

    fields = body.get("tespit_edilen_alanlar", {})
    controls = body.get("periodic_controls_ekle", [])
    file_info = body.get("file_info", {})
    belge_turu = body.get("belge_turu", "MEBBİS İSG Belgesi")

    conn = get_db()
    c = conn.cursor()

    # 1. Okul kontrolü
    c.execute("SELECT id, kurum_kodu, kurum_adi FROM schools WHERE id = ?", (school_id,))
    school = c.fetchone()
    if not school:
        conn.close()
        raise HTTPException(status_code=404, detail="Okul bulunamadı.")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 2. schools tablosu güncellemeleri
    school_updates = []
    school_params = []
    for k in ["ogrenci_sayisi", "ogretmen_sayisi", "mudur_ad_soyad", "mudur_tel"]:
        val = fields.get(k)
        if val is not None and str(val).strip():
            school_updates.append(f"{k} = ?")
            school_params.append(val)

    user_name = user["full_name"] if (user and "full_name" in user.keys()) else "MEBBİS AI"

    if school_updates:
        school_updates.append("son_veri_giris_tarihi = ?")
        school_updates.append("son_veri_giren_kisi = ?")
        school_params.extend([now_str, f"Yapay Zeka ({user_name})", school_id])
        c.execute(f"UPDATE schools SET {', '.join(school_updates)} WHERE id = ?", tuple(school_params))
    else:
        c.execute("UPDATE schools SET son_veri_giris_tarihi = ?, son_veri_giren_kisi = ? WHERE id = ?", (now_str, f"Yapay Zeka ({user_name})", school_id))

    # 3. inspection_metrics tablosu güncellemeleri
    metric_cols = [
        "kazan_kontrol_tarihi", "kazan_durumu",
        "elektrik_topraklama_tarihi",
        "paratoner_kontrol_tarihi", "paratoner_durumu",
        "yangin_tupleri_kontrol_tarihi",
        "asansor_muayene_tarihi", "asansor_durumu",
        "su_deposu_temizlik_tarihi", "su_deposu_durumu",
        "adep_var_mi", "adep_tarihi",
        "tatbikat_tarihi", "tatbikat_turu", "tatbikat_katilimci_sayisi",
        "ilkyardimci_personel_sayisi", "ozel_guvenlik_sayisi",
        "toplam_kamera_sayisi"
    ]
    metric_updates = []
    metric_params = []
    for col in metric_cols:
        val = fields.get(col)
        if val is not None and str(val).strip():
            metric_updates.append(f"{col} = ?")
            metric_params.append(val)

    if metric_updates:
        metric_updates.append("updated_at = ?")
        metric_params.extend([now_str, school_id])
        c.execute(f"UPDATE inspection_metrics SET {', '.join(metric_updates)} WHERE school_id = ?", tuple(metric_params))

    # 4. periodic_controls eklemeleri
    added_controls_count = 0
    for ctrl in controls:
        ekipman = ctrl.get("ekipman_turu")
        tarih = ctrl.get("son_kontrol_tarihi")
        if ekipman and tarih:
            gecerlilik = int(ctrl.get("gecerlilik_ayi") or 12)
            sonraki = ctrl.get("sonraki_kontrol_tarihi")
            if not sonraki:
                try:
                    dt = datetime.strptime(tarih, "%Y-%m-%d")
                    sonraki = (dt + timedelta(days=gecerlilik * 30)).strftime("%Y-%m-%d")
                except Exception:
                    sonraki = tarih
            kurum = ctrl.get("kontrol_yapan_kurum") or "MEBBİS Yetkili Muayene Kuruluşu"
            sonuc = ctrl.get("rapor_sonucu") or "Gecerli"
            notlar = ctrl.get("notlar") or "Yapay Zeka (MEBBİS AI Aktarımı) ile otomatik işlendi."
            c.execute("""
                INSERT INTO periodic_controls (school_id, ekipman_turu, son_kontrol_tarihi, gecerlilik_ayi, sonraki_kontrol_tarihi, kontrol_yapan_kurum, rapor_sonucu, alarm_seviyesi, notlar)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'Gecerli', ?)
            """, (school_id, ekipman, tarih, gecerlilik, sonraki, kurum, sonuc, notlar))
            added_controls_count += 1

    # 5. documents_archive kaydı & Akıllı Hiyerarşik Depolama
    final_file_url = None
    storage_savings = None
    if file_info and file_info.get("file_path"):
        raw_rel = file_info.get("file_path").lstrip("/\\")
        temp_abs = os.path.join(BASE_DIR, raw_rel)
        orig_name = file_info.get("filename", "MEBBİS_Belgesi")
        if os.path.exists(temp_abs):
            try:
                with open(temp_abs, "rb") as tf:
                    raw_bytes = tf.read()
                # Akıllı depolama ve görsel sıkıştırma motoru
                saved_res = isg_storage.save_school_document(
                    file_bytes=raw_bytes,
                    original_filename=orig_name,
                    school_info=dict(school),
                    belge_turu=belge_turu
                )
                final_file_url = saved_res["relative_url"]
                storage_savings = saved_res
                try:
                    os.remove(temp_abs)
                except Exception:
                    pass
            except Exception as e:
                final_file_url = file_info.get("file_path")
        else:
            final_file_url = file_info.get("file_path")

        c.execute("""
            INSERT INTO documents_archive (school_id, belge_turu, belge_adi, dosya_yolu, yukleyen_user_id, yukleme_tarihi)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (school_id, belge_turu, orig_name, final_file_url or file_info.get("file_path"), user["id"], now_str))

    conn.commit()
    conn.close()

    return {
        "success": True,
        "message": f"{school['kurum_adi']} okulunun verileri ve periyodik kontrolleri başarıyla güncellendi.",
        "updated_school": school["kurum_adi"],
        "added_controls": added_controls_count,
        "storage_savings": storage_savings
    }

@app.get("/admin/arsiv-indir/{year}")
def download_annual_archive(year: int, request: Request):
    """Belirli bir yıla ait tüm okul belgelerini tek bir zip dosyası olarak indirir."""
    user = get_current_user(request)
    if not user or user["role"] not in ["koordinator", "uzman", "admin", "isg_uzmani"]:
        raise HTTPException(status_code=403, detail="Yalnızca koordinatör ve İSG uzmanları yıllık arşiv indirebilir.")

    relative_url, zip_abs_path, zip_size_mb = isg_storage.export_annual_archive_zip(year)
    if not os.path.exists(zip_abs_path):
        raise HTTPException(status_code=404, detail=f"{year} yılına ait arşiv bulunamadı.")

    return FileResponse(
        path=zip_abs_path,
        filename=f"muratpasa_isg_arsiv_{year}.zip",
        media_type="application/zip"
    )

@app.get("/api/storage/stats")
def api_storage_stats(request: Request):
    """Canlı depolama ve arşivleme istatistiklerini döner."""
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Oturum açmanız gerekiyor.")
    return isg_storage.get_storage_statistics()

# MODÜL 4: OKULUM TEMİZ DIŞ TETKİK VE BELGELENDİRME PORTALI
@app.get("/okulum-temiz", response_class=HTMLResponse)
def okulum_temiz_page(request: Request, auditor: Optional[str] = None):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    conn = get_db()
    c = conn.cursor()
    query = """
        SELECT a.*, s.kurum_adi, s.kurum_kodu, s.kurum_turu, s.kademe, s.mudur_ad_soyad, s.mudur_tel
        FROM okulum_temiz_audits a
        JOIN schools s ON a.school_id = s.id
    """
    params = []
    if auditor:
        query += " WHERE a.tetkikciler LIKE ?"
        params.append(f"%{auditor}%")
    query += " ORDER BY a.tetkik_tarihi ASC, a.id ASC"
    c.execute(query, params)
    audits = c.fetchall()

    c.execute("SELECT count(*) FROM okulum_temiz_audits")
    tot = c.fetchone()[0]
    c.execute("SELECT count(*) FROM okulum_temiz_audits WHERE tetkikciler LIKE '%Ramazan ERYILDIZ%'")
    eryildiz_tot = c.fetchone()[0]
    c.execute("SELECT count(*) FROM okulum_temiz_audits WHERE sonuc = 'SARTLI_UYGUN_DOF' OR dof_sayisi > 0")
    dof_cnt = c.fetchone()[0]

    c.execute("SELECT id, kurum_kodu, kurum_adi FROM schools WHERE is_active_school = 1 ORDER BY kurum_adi ASC")
    all_schools = c.fetchall()
    conn.close()

    return templates.TemplateResponse(request=request, name="okulum_temiz.html", context={
        "user": user,
        "audits": audits,
        "total_audits": tot,
        "eryildiz_audits": eryildiz_tot,
        "dof_count": dof_cnt,
        "filter_auditor": auditor,
        "all_schools": all_schools,
        "active_page": "okulum_temiz"
    })

@app.post("/okulum-temiz/ekle")
def add_okulum_temiz_audit(school_id: int = Form(...), tetkik_tarihi: str = Form(...), tetkik_saati: str = Form("09:30 - 12:30"), tetkikciler: str = Form(...), tetkik_turu: str = Form("Gözetim")):
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        INSERT INTO okulum_temiz_audits (school_id, tetkik_tarihi, tetkik_saati, tetkikciler, tetkik_turu, sonuc)
        VALUES (?, ?, ?, ?, ?, 'Planlandı')
    """, (school_id, tetkik_tarihi, tetkik_saati, tetkikciler, tetkik_turu))
    conn.commit()
    conn.close()
    return RedirectResponse("/okulum-temiz", status_code=303)

@app.get("/okulum-temiz/form/{audit_id}", response_class=HTMLResponse)
def okulum_temiz_form_page(request: Request, audit_id: int):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT a.*, s.kurum_adi, s.kurum_kodu, s.kurum_turu, s.kademe, s.mudur_ad_soyad, s.mudur_tel
        FROM okulum_temiz_audits a
        JOIN schools s ON a.school_id = s.id
        WHERE a.id = ?
    """, (audit_id,))
    audit = c.fetchone()
    conn.close()
    if not audit:
        raise HTTPException(status_code=404, detail="Tetkik bulunamadı.")

    results = {}
    if audit['kriter_sonuclari']:
        try:
            results = json.loads(audit['kriter_sonuclari'])
        except Exception:
            results = {}

    return templates.TemplateResponse(request=request, name="okulum_temiz_form.html", context={
        "user": user,
        "audit": audit,
        "sections": OKULUM_TEMIZ_SECTIONS,
        "results": results,
        "active_page": "okulum_temiz"
    })

@app.post("/okulum-temiz/kaydet/{audit_id}")
async def save_okulum_temiz_form(request: Request, audit_id: int):
    form_data = await request.form()
    results = {}
    dof_items = []
    dof_count = 0

    flat_q = get_all_questions_flat()
    for q in flat_q:
        q_no = q["q_no"]
        status = form_data.get(f"q_{q_no}", "Uygun")
        note = form_data.get(f"note_{q_no}", "")
        results[q_no] = {"status": status, "note": note}
        if status == "Uygun Değil":
            dof_count += 1
            dof_items.append({"no": q_no, "soru": q["q_text"], "bulgu": note})

    sonuc = form_data.get("sonuc", "BELGE_VERILMESI_UYGUN")
    dof_termini = form_data.get("dof_termini", "15 Gün")
    notlar = form_data.get("notlar", "")

    conn = get_db()
    c = conn.cursor()
    c.execute("""
        UPDATE okulum_temiz_audits
        SET kriter_sonuclari = ?,
            dof_sayisi = ?,
            dof_maddeleri = ?,
            dof_termini = ?,
            sonuc = ?,
            notlar = ?
        WHERE id = ?
    """, (
        json.dumps(results, ensure_ascii=False),
        dof_count,
        json.dumps(dof_items, ensure_ascii=False),
        dof_termini,
        sonuc,
        notlar,
        audit_id
    ))
    conn.commit()
    conn.close()
    return RedirectResponse("/okulum-temiz", status_code=303)

# MODÜL 5: DÖF VE MUAYENE TAKİP MOTORU
@app.get("/dof-takip", response_class=HTMLResponse)
def dof_tracking_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT s.kurum_adi, s.kurum_kodu, s.mudur_ad_soyad, s.mudur_tel, COUNT(r.id) as open_dof
        FROM risk_assessments r
        JOIN schools s ON r.school_id = s.id
        WHERE r.dof_durumu != 'Kapatıldı' AND s.is_active_school = 1
        GROUP BY s.id
        ORDER BY open_dof DESC
        LIMIT 10
    """)
    top_dof = c.fetchall()
    c.execute("SELECT COUNT(*) FROM risk_assessments WHERE dof_durumu != 'Kapatıldı'")
    tot_open = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM periodic_controls WHERE alarm_seviyesi = 'Yaklasiyor (30 Gun)'")
    upc = c.fetchone()[0]
    conn.close()

    return templates.TemplateResponse(request=request, name="dof_tracking.html", context={
        "user": user,
        "top_dof_schools": top_dof,
        "total_open_dof": tot_open,
        "expired_dof_count": 187,
        "upcoming_controls": upc,
        "active_page": "dof_takip"
    })

# MODÜL 6: KURUMSAL İSG DOKÜMAN KÜTÜPHANESİ
@app.get("/dokuman-kutuphanesi", response_class=HTMLResponse)
def document_library_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    return templates.TemplateResponse(request=request, name="document_library.html", context={
        "user": user,
        "active_page": "library"
    })

# ==================== KULLANICI YÖNETİMİ (YALNIZCA KOORDİNATÖR) ====================

@app.get("/kullanicilar", response_class=HTMLResponse)
def users_management_page(request: Request, msg: Optional[str] = None, error: Optional[str] = None):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    if user["role"] != "koordinator":
        raise HTTPException(status_code=403, detail="Bu sayfaya sadece İlçe İSG Büro Koordinatörü erişebilir.")

    conn = get_db()
    c = conn.cursor()
    
    # Kullanıcıları okul zimmet sayılarıyla birlikte getir
    c.execute("""
        SELECT u.*,
               COUNT(s.id) as zimmetli_okul_sayisi
        FROM users u
        LEFT JOIN schools s ON u.id = s.sorumlu_uzman_id AND s.is_active_school = 1
        GROUP BY u.id
        ORDER BY CASE WHEN u.role = 'koordinator' THEN 1 ELSE 2 END, u.id ASC
    """)
    all_users = [dict(r) for r in c.fetchall()]

    # İstatistikler
    c.execute("SELECT COUNT(*) FROM users")
    total_users = c.fetchone()[0]
    
    c.execute("SELECT COUNT(*) FROM users WHERE role = 'koordinator'")
    coord_count = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM users WHERE role = 'uzman'")
    uzman_count = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM users WHERE igu_sinifi LIKE '%A%'")
    a_igu_count = c.fetchone()[0]

    conn.close()

    db_size = 0.0
    if os.path.exists(DB_PATH):
        db_size = round(os.path.getsize(DB_PATH) / (1024 * 1024), 2)

    return templates.TemplateResponse(request=request, name="users_management.html", context={
        "user": user,
        "all_users": all_users,
        "stats": {
            "total_users": total_users,
            "coord_count": coord_count,
            "uzman_count": uzman_count,
            "a_igu_count": a_igu_count
        },
        "db_size_mb": db_size,
        "gdrive_webhook_url": os.environ.get("GDRIVE_BACKUP_WEBHOOK_URL", ""),
        "backup_secret_token": os.environ.get("BACKUP_SECRET_TOKEN", "muratpasa-isg-yedek-2026"),
        "msg": msg,
        "error": error,
        "active_page": "users"
    })

@app.post("/kullanicilar/ekle")
def add_user(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    full_name: str = Form(...),
    title: str = Form("İş Güvenliği Uzmanı"),
    role: str = Form("uzman"),
    email: str = Form(""),
    phone: str = Form(""),
    igu_sinifi: str = Form("C Sınıfı İGU")
):
    user = get_current_user(request)
    if not user or user["role"] != "koordinator":
        raise HTTPException(status_code=403, detail="Yetkisiz işlem.")

    clean_username = username.strip().lower()
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id FROM users WHERE username = ?", (clean_username,))
    existing = c.fetchone()
    if existing:
        conn.close()
        return RedirectResponse("/kullanicilar?error=username_exists", status_code=303)

    p_hash = hash_pw(password.strip())
    c.execute("""
        INSERT INTO users (username, password_hash, full_name, title, role, email, phone, igu_sinifi)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (clean_username, p_hash, full_name.strip(), title.strip(), role, email.strip(), phone.strip(), igu_sinifi.strip()))
    conn.commit()
    conn.close()

    return RedirectResponse("/kullanicilar?msg=added", status_code=303)

@app.post("/kullanicilar/guncelle/{target_id}")
def update_user(
    request: Request,
    target_id: int,
    full_name: str = Form(...),
    title: str = Form(...),
    role: str = Form(...),
    email: str = Form(""),
    phone: str = Form(""),
    igu_sinifi: str = Form(""),
    new_password: Optional[str] = Form(None)
):
    user = get_current_user(request)
    if not user or user["role"] != "koordinator":
        raise HTTPException(status_code=403, detail="Yetkisiz işlem.")

    conn = get_db()
    c = conn.cursor()

    if new_password and new_password.strip():
        p_hash = hash_pw(new_password.strip())
        c.execute("""
            UPDATE users SET
                full_name = ?, title = ?, role = ?, email = ?, phone = ?, igu_sinifi = ?, password_hash = ?
            WHERE id = ?
        """, (full_name.strip(), title.strip(), role, email.strip(), phone.strip(), igu_sinifi.strip(), p_hash, target_id))
    else:
        c.execute("""
            UPDATE users SET
                full_name = ?, title = ?, role = ?, email = ?, phone = ?, igu_sinifi = ?
            WHERE id = ?
        """, (full_name.strip(), title.strip(), role, email.strip(), phone.strip(), igu_sinifi.strip(), target_id))

    conn.commit()
    conn.close()

    return RedirectResponse("/kullanicilar?msg=updated", status_code=303)

@app.post("/kullanicilar/sil/{target_id}")
def delete_user(request: Request, target_id: int):
    user = get_current_user(request)
    if not user or user["role"] != "koordinator":
        raise HTTPException(status_code=403, detail="Yetkisiz işlem.")

    if user["id"] == target_id:
        return RedirectResponse("/kullanicilar?error=cannot_delete_self", status_code=303)

    conn = get_db()
    c = conn.cursor()
    # Okulları boşa düşür (yetkisiz kalmaması için)
    c.execute("UPDATE schools SET sorumlu_uzman_id = NULL WHERE sorumlu_uzman_id = ?", (target_id,))
    c.execute("DELETE FROM users WHERE id = ?", (target_id,))
    conn.commit()
    conn.close()

    return RedirectResponse("/kullanicilar?msg=deleted", status_code=303)

# ==================== VERİTABANI YEDEKLEME & AFET KURTARMA ====================

@app.get("/admin/veritabani-indir")
def download_database_backup(request: Request, token: Optional[str] = None):
    user = get_current_user(request)
    secret_token = os.environ.get("BACKUP_SECRET_TOKEN", "muratpasa-isg-yedek-2026")
    is_authorized = (user and user["role"] == "koordinator") or (token and token.strip() == secret_token)
    if not is_authorized:
        raise HTTPException(status_code=403, detail="Yalnızca İlçe İSG Büro Koordinatörü canlı veritabanı yedeği indirebilir.")

    if not os.path.exists(DB_PATH):
        raise HTTPException(status_code=404, detail="Veritabanı dosyası bulunamadı.")

    filename = f"isg_muratpasa_canli_yedek_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    return FileResponse(
        path=DB_PATH,
        filename=filename,
        media_type="application/x-sqlite3"
    )

@app.post("/admin/veritabani-yukle")
async def restore_database_backup(request: Request, backup_file: UploadFile = File(...)):
    user = get_current_user(request)
    if not user or user["role"] != "koordinator":
        raise HTTPException(status_code=403, detail="Yetkisiz işlem.")

    content = await backup_file.read()
    # SQLite başlık doğrulaması (İlk 16 bayt: 'SQLite format 3\x00')
    if not content.startswith(b"SQLite format 3\x00"):
        return RedirectResponse("/kullanicilar?error=invalid_db_file", status_code=303)

    # Mevcut veritabanını güvenle üzerine yaz
    with open(DB_PATH, "wb") as f:
        f.write(content)

    return RedirectResponse("/kullanicilar?msg=db_restored", status_code=303)

@app.post("/admin/drive-yedekle")
def trigger_drive_backup(request: Request, webhook_url: Optional[str] = Form(None)):
    user = get_current_user(request)
    if not user or user["role"] != "koordinator":
        raise HTTPException(status_code=403, detail="Yetkisiz işlem.")

    target_url = (webhook_url or os.environ.get("GDRIVE_BACKUP_WEBHOOK_URL") or "").strip()
    if not target_url:
        return RedirectResponse("/kullanicilar?error=no_webhook_url", status_code=303)

    try:
        import base64
        with open(DB_PATH, "rb") as f:
            db_bytes = f.read()

        b64_data = base64.b64encode(db_bytes).decode("utf-8")
        resp = requests.post(
            target_url,
            json={
                "filename": f"isg_muratpasa_otomatik_yedek_{datetime.now().strftime('%Y%m%d_%H%M')}.db",
                "file_data": b64_data
            },
            headers={"Content-Type": "application/json"},
            allow_redirects=True,
            timeout=60
        )
        if resp.status_code in [200, 201, 302]:
            return RedirectResponse("/kullanicilar?msg=drive_backup_success", status_code=303)
        else:
            print(f"Drive backup returned status: {resp.status_code}, response: {resp.text[:200]}")
            return RedirectResponse("/kullanicilar?error=drive_backup_failed", status_code=303)
    except Exception as e:
        print(f"Drive backup exception: {e}")
        return RedirectResponse("/kullanicilar?error=drive_backup_failed", status_code=303)

# ==================== MODÜL 7: KULLANICI GERİ BİLDİRİM, İSTEK HAVUZU VE AI GELİŞTİRİCİ ASİSTANI ====================

@app.post("/api/feedback/submit")
async def api_submit_feedback(request: Request):
    """Kullanıcının ilettiği geliştirme/şablon isteğini Gemini AI ile analiz eder ve veritabanına kaydeder."""
    body = await request.json()
    kurum_kodu = (body.get("kurum_kodu") or "").strip()
    gonderen_ad_soyad = (body.get("gonderen_ad_soyad") or "").strip()
    unvan = (body.get("unvan") or "").strip()
    iletisim_bilgisi = (body.get("iletisim_bilgisi") or "").strip()
    istek_turu = body.get("istek_turu", "DIGER")
    talep_basligi = (body.get("talep_basligi") or "").strip()
    detayli_aciklama = (body.get("detayli_aciklama") or "").strip()
    ornek_yapi = (body.get("ornek_dokuman_veya_form_yapisi") or "").strip()

    if not talep_basligi or not detayli_aciklama:
        raise HTTPException(status_code=400, detail="Talep başlığı ve detaylı açıklama zorunludur.")

    conn = get_db()
    c = conn.cursor()

    school_id = None
    school_name = ""
    if kurum_kodu:
        c.execute("SELECT id, kurum_adi FROM schools WHERE kurum_kodu = ?", (kurum_kodu,))
        sc = c.fetchone()
        if sc:
            school_id = sc["id"]
            school_name = sc["kurum_adi"]

    # Gemini AI ile Analiz & Task Üretimi
    ai_analysis = isg_ai.analyze_feedback_request(
        gonderen_ad_soyad=gonderen_ad_soyad,
        unvan=unvan,
        okul_adi=school_name or kurum_kodu,
        talep_basligi=talep_basligi,
        detayli_aciklama=detayli_aciklama,
        ornek_yapi=ornek_yapi
    )

    ai_prompt_text = json.dumps(ai_analysis, ensure_ascii=False, indent=2)
    ai_cat = ai_analysis.get("oneri_kategori", istek_turu)

    c.execute("""
        INSERT INTO feature_requests (
            kurum_kodu, school_id, gonderen_ad_soyad, unvan, iletisim_bilgisi,
            istek_turu, talep_basligi, detayli_aciklama, ornek_dokuman_veya_form_yapisi,
            ai_asistan_ozeti_ve_prompt, ai_oneri_kategori, durum
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'AI_ANALIZ_EDILDI')
    """, (
        kurum_kodu, school_id, gonderen_ad_soyad, unvan, iletisim_bilgisi,
        istek_turu, talep_basligi, detayli_aciklama, ornek_yapi,
        ai_prompt_text, ai_cat
    ))

    new_id = c.lastrowid
    conn.commit()
    conn.close()

    return {
        "success": True,
        "message": "Talebiniz başarıyla iletildi. Yapay zeka asistanı teknik analizi tamamlayıp geliştirici iş listesine ekledi.",
        "request_id": new_id,
        "ai_analysis": ai_analysis
    }

@app.get("/istek-havuzu", response_class=HTMLResponse)
def feedback_hub_page(request: Request, status: Optional[str] = None):
    """Geliştirici Yönetim Paneli (Ramazan ERYILDIZ Dev Dashboard)."""
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    c = conn.cursor()

    query = """
        SELECT f.*, s.kurum_adi 
        FROM feature_requests f
        LEFT JOIN schools s ON f.school_id = s.id
    """
    params = []
    if status:
        query += " WHERE f.durum = ?"
        params.append(status)
    query += " ORDER BY f.id DESC"

    c.execute(query, tuple(params))
    requests_list = [dict(r) for r in c.fetchall()]

    c.execute("SELECT COUNT(*) FROM feature_requests")
    total_reqs = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM feature_requests WHERE durum IN ('BEKLEMEDE', 'AI_ANALIZ_EDILDI')")
    pending_reqs = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM feature_requests WHERE durum = 'GELISMEDE'")
    in_progress = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM feature_requests WHERE durum = 'YAYINLANDI'")
    completed = c.fetchone()[0]

    conn.close()

    return templates.TemplateResponse(request=request, name="feedback_hub.html", context={
        "user": user,
        "requests": requests_list,
        "total_reqs": total_reqs,
        "pending_reqs": pending_reqs,
        "in_progress": in_progress,
        "completed": completed,
        "current_filter": status,
        "active_page": "feedback_hub"
    })

@app.post("/api/feedback/update")
async def api_update_feedback(request: Request):
    """Geliştirici tarafından durum ve not güncellemesi."""
    user = get_current_user(request)
    if not user or user["role"] not in ["koordinator", "uzman", "admin"]:
        raise HTTPException(status_code=403, detail="Yetkisiz işlem.")

    body = await request.json()
    req_id = body.get("id")
    durum = body.get("durum")
    gelistirici_notu = body.get("gelistirici_notu")

    if not req_id or not durum:
        raise HTTPException(status_code=400, detail="Eksik parametre.")

    conn = get_db()
    c = conn.cursor()
    c.execute("""
        UPDATE feature_requests 
        SET durum = ?, gelistirici_notu = ?, updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
    """, (durum, gelistirici_notu, req_id))
    conn.commit()
    conn.close()

    return {"success": True, "message": "Talep durumu güncellendi."}

# ==================== MODÜL 8: OKULA, ATÖLYEYE VE BİRİME ÖZGÜ DİNAMİK RİSK DEĞERLENDİRME & MEBBİS KÖPRÜLEME MOTORU ====================

MTAL_WORKSHOP_MODULES = [
    ("MAKINE_CNC", "Makine ve Tasarım Teknolojisi (Torna, Freze, CNC, Taşlama)"),
    ("ELEKTRIK_ELEKTRONIK", "Elektrik-Elektronik Teknolojisi (Pano, Bobinaj, Yüksek Gerilim)"),
    ("MOTORLU_ARACLAR", "Motorlu Araçlar Teknolojisi (Lift, Şasi, Fren Test, Akü Şarj)"),
    ("METAL_KAYNAK", "Metal Teknolojisi (Elektrik Ark, Gazaltı / TIG Kaynağı, Oksigaz)"),
    ("TESISAT_IKLIMLENDIRME", "Tesisat Teknolojisi ve İklimlendirme (Kombi, Gaz Tesisatı, Brülör)"),
    ("BILISIM_AG", "Bilişim Teknolojileri ve Ağ İşletmenliği (Server, Kablaj)"),
    ("YIYECEK_ICECEK", "Yiyecek İçecek Hizmetleri (Mutfak, Endüstriyel Fırın, Gazlı Ocaklar)"),
    ("KIMYA_TEKNOLOJISI", "Kimya Teknolojisi (Çeker Ocak, Asit-Baz, Kimyasal Depo)"),
    ("MOBILYA_IC_MEKAN", "Mobilya ve İç Mekan Tasarımı (Daire Testere, Planya)"),
    ("KONAKLAMA_SEYAHAT", "Konaklama ve Seyahat Hizmetleri (Kat Hizmetleri, Çamaşırhane)"),
    ("GUZELLIK_SAC", "Güzellik ve Saç Bakım Hizmetleri (Kimyasal Boyalar, Cihazlar)"),
    ("ENDUSTRIYEL_OTOMASYON", "Endüstriyel Otomasyon (PLC, Robot Kollar, Pnömatik)")
]

@app.get("/dinamik-risk-degerlendirme", response_class=HTMLResponse)
def dynamic_risk_assessment_page(request: Request, school_id: Optional[int] = None):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, kurum_kodu, kurum_adi, kademe, tehlike_sinifi FROM schools WHERE is_active_school = 1 ORDER BY kurum_adi ASC")
    schools = [dict(r) for r in c.fetchall()]

    selected_school = None
    school_risks = []
    if school_id:
        c.execute("SELECT * FROM schools WHERE id = ?", (school_id,))
        s_row = c.fetchone()
        if s_row:
            selected_school = dict(s_row)
            c.execute("""
                SELECT * FROM risk_assessments 
                WHERE school_id = ? 
                ORDER BY risk_skoru DESC, id DESC
            """, (school_id,))
            school_risks = [dict(r) for r in c.fetchall()]

    conn.close()

    return templates.TemplateResponse(request=request, name="dynamic_risk.html", context={
        "user": user,
        "schools": schools,
        "selected_school": selected_school,
        "school_risks": school_risks,
        "workshop_modules": MTAL_WORKSHOP_MODULES,
        "active_page": "dynamic_risk"
    })

@app.post("/api/risk/evaluate-and-bridge")
async def api_evaluate_and_bridge(request: Request):
    """Fine-Kinney veya 5x5 skorunu hesaplar ve MEBBİS köprüleme sihirbazını çalıştırır."""
    body = await request.json()
    method = body.get("degerlendirme_metodu", "FINE_KINNEY")
    olasilik = float(body.get("olasilik") or 1.0)
    frekans = float(body.get("frekans") or 1.0)
    siddet = float(body.get("siddet") or 1.0)

    bolum_unite = body.get("bolum_unite", "BINA_GENEL")
    meslek_alani = body.get("meslek_alani_veya_atolye_adi", "")
    tehlike_kaynagi = body.get("tehlike_kaynagi", "")
    risk_senaryosu = body.get("olasi_tehlike_ve_risk_senaryosu", "")
    mevcut_onlemler = body.get("mevcut_onlemler", "")
    dof_onlem = body.get("alinmasi_gereken_dof_onlem", "")

    # Hesaplama
    if method == "FINE_KINNEY":
        risk_skoru = round(olasilik * frekans * siddet, 1)
        if risk_skoru < 20:
            risk_seviyesi = "DUSUK_KABUL_EDILEBILIR"
            seviye_etiket = "Kabul Edilebilir Risk (< 20)"
            seviye_renk = "emerald"
        elif risk_skoru < 70:
            risk_seviyesi = "ORTA_OLASI"
            seviye_etiket = "Olası Risk / Gözetim Altında Tutulmalı (20 - 70)"
            seviye_renk = "amber"
        elif risk_skoru < 200:
            risk_seviyesi = "YUKSEK_ONEMLI"
            seviye_etiket = "Önemli Risk / DÖF Başlatılmalı (70 - 200)"
            seviye_renk = "orange"
        elif risk_skoru < 400:
            risk_seviyesi = "YUKSEK_ACIL"
            seviye_etiket = "Yüksek Risk / Acil Tedbir Alınmalı (200 - 400)"
            seviye_renk = "rose"
        else:
            risk_seviyesi = "COK_YUKSEK_TOLERANS_EDILEMEZ"
            seviye_etiket = "Tolerans Edilemez / İşi Durdur (>= 400)"
            seviye_renk = "red"
    else:  # 5x5 L Tipi
        risk_skoru = round(olasilik * siddet, 1)
        if risk_skoru <= 5:
            risk_seviyesi = "DUSUK_KABUL_EDILEBILIR"
            seviye_etiket = "Düşük Risk (1 - 5)"
            seviye_renk = "emerald"
        elif risk_skoru <= 12:
            risk_seviyesi = "ORTA_OLASI"
            seviye_etiket = "Orta Risk (6 - 12)"
            seviye_renk = "amber"
        elif risk_skoru < 20:
            risk_seviyesi = "YUKSEK_ONEMLI"
            seviye_etiket = "Yüksek Risk (15 - 19)"
            seviye_renk = "orange"
        else:
            risk_seviyesi = "COK_YUKSEK_TOLERANS_EDILEMEZ"
            seviye_etiket = "Çok Yüksek Risk (20 - 25)"
            seviye_renk = "red"

    # MEBBİS Köprüleme Sihirbazı
    bridge_data = isg_ai.bridge_mebbis_risk(
        bolum_unite=bolum_unite,
        meslek_alani_veya_atolye=meslek_alani,
        tehlike_kaynagi=tehlike_kaynagi,
        risk_senaryosu=risk_senaryosu,
        mevcut_onlemler=mevcut_onlemler,
        dof_onlem=dof_onlem,
        risk_skoru=risk_skoru,
        risk_seviyesi=seviye_etiket
    )

    return {
        "success": True,
        "risk_skoru": risk_skoru,
        "risk_seviyesi": risk_seviyesi,
        "seviye_etiket": seviye_etiket,
        "seviye_renk": seviye_renk,
        "bridge": bridge_data
    }

@app.post("/api/risk/save")
async def api_save_risk(request: Request):
    """Yeni risk değerlendirme kaydını veritabanına ekler."""
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Oturum açmanız gerekiyor.")

    body = await request.json()
    school_id = body.get("school_id")
    if not school_id:
        raise HTTPException(status_code=400, detail="Okul seçilmedi.")

    conn = get_db()
    c = conn.cursor()

    c.execute("""
        INSERT INTO risk_assessments (
            school_id, kontrol_edilen_birim, bolum_unite, meslek_alani_veya_atolye_adi,
            tehlike, tehlike_unsurlari, riskler,
            degerlendirme_metodu, olasilik, frekans, siddet, risk_skoru, risk_puani, risk_seviyesi,
            alinan_onlemler, alinmasi_gereken_dof_onlem,
            dayanak_mevzuat, mebbis_kutuphane_kodu, mebbis_dof_aciklama_metni,
            odenek_talep_kapsaminda_mi, termin_tarihi, sorumlular, dof_durumu,
            denetleyen_uzman, ekip_onayi
        ) VALUES (
            ?, ?, ?, ?,
            ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?,
            ?, ?,
            ?, ?, ?,
            ?, ?, ?, 'Açık DÖF',
            ?, 1
        )
    """, (
        school_id,
        body.get("kontrol_edilen_birim") or body.get("bolum_unite") or "Genel Bina",
        body.get("bolum_unite"),
        body.get("meslek_alani_veya_atolye_adi"),
        body.get("tehlike_kaynagi"),
        body.get("olasi_tehlike_ve_risk_senaryosu"),
        body.get("olasi_tehlike_ve_risk_senaryosu"),
        body.get("degerlendirme_metodu", "FINE_KINNEY"),
        float(body.get("olasilik") or 1.0),
        float(body.get("frekans") or 1.0),
        float(body.get("siddet") or 1.0),
        float(body.get("risk_skoru") or 0.0),
        float(body.get("risk_skoru") or 0.0),
        body.get("risk_seviyesi"),
        body.get("mevcut_onlemler"),
        body.get("alinmasi_gereken_dof_onlem"),
        body.get("dayanak_mevzuat"),
        body.get("mebbis_kutuphane_kodu"),
        body.get("mebbis_dof_aciklama_metni"),
        1 if body.get("odenek_talep_kapsaminda_mi") else 0,
        body.get("termin_tarihi"),
        body.get("sorumlular") or "İşveren Vekili (Müdür Yrd.) / İSG Çalışan Temsilcisi",
        user["full_name"]
    ))

    new_id = c.lastrowid
    conn.commit()
    conn.close()

    return {
        "success": True,
        "message": "Risk değerlendirmesi başarıyla kaydedildi.",
        "risk_id": new_id
    }

@app.get("/risk-analizi-raporu/{school_id}", response_class=HTMLResponse)
def print_risk_report_page(request: Request, school_id: int):
    """6331 sayılı Kanun ve Risk Değerlendirmesi Yönetmeliğine tam uyumlu resmî çıktı ekranı."""
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM schools WHERE id = ?", (school_id,))
    school = c.fetchone()
    if not school:
        conn.close()
        raise HTTPException(status_code=404, detail="Okul bulunamadı.")

    c.execute("""
        SELECT * FROM risk_assessments 
        WHERE school_id = ? 
        ORDER BY risk_skoru DESC, id ASC
    """, (school_id,))
    risks = [dict(r) for r in c.fetchall()]
    conn.close()

    return templates.TemplateResponse(request=request, name="risk_report_print.html", context={
        "user": user,
        "school": dict(school),
        "risks": risks,
        "expert_name": "Ramazan ERYILDIZ",
        "expert_title": "A Sınıfı İş Güvenliği Uzmanı (İlçe MEM İSG Büro)",
        "today": datetime.now().strftime("%d.%m.%Y")
    })

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)

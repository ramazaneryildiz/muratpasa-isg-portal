"""
Muratpaşa İlçe MEM İSG Bilgi Yönetim Portalı
Google Gemini Yapay Zeka Entegrasyon Motoru (isg_ai.py)
Powered by Google AI Studio - Gemini 3.5 Flash
"""

import os
import json
import logging
import requests

logger = logging.getLogger("isg_ai")

# Ortam değişkenlerinden API anahtarını al (yoksa .env dosyasını oku)
def get_api_key():
    key = os.environ.get("GEMINI_API_KEY")
    if not key:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        env_file = os.path.join(base_dir, ".env")
        if os.path.exists(env_file):
            with open(env_file, "r", encoding="utf-8") as f:
                for line in f:
                    if "=" in line and not line.strip().startswith("#"):
                        k, v = line.strip().split("=", 1)
                        if k.strip() == "GEMINI_API_KEY":
                            key = v.strip()
                            os.environ["GEMINI_API_KEY"] = key
                            break
    return key or ""

CANDIDATE_MODELS = [
    os.environ.get("GEMINI_MODEL", "gemini-3.5-flash"),
    "gemini-3.5-flash-lite",
    "gemini-flash-latest",
    "gemini-3.1-flash-lite"
]

def call_gemini(prompt: str, json_mode: bool = False, timeout: int = 35) -> str:
    """Gemini API'sini çağırır. Model meşgulse (503/404) otomatik olarak yedek modele geçer."""
    parts = [{"text": prompt}]
    return call_gemini_multimodal(parts, json_mode=json_mode, timeout=timeout)

def call_gemini_multimodal(parts: list, json_mode: bool = False, timeout: int = 45) -> str:
    """Gemini API'sini metin ve görsellerle (Multimodal) çağırır."""
    import time
    api_key = get_api_key()
    if not api_key:
        raise ValueError("GEMINI_API_KEY bulunamadı.")

    payload = {
        "contents": [{"parts": parts}]
    }
    if json_mode:
        payload["generationConfig"] = {"response_mime_type": "application/json"}

    last_error = None
    for model_name in CANDIDATE_MODELS:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        for attempt in range(2):
            try:
                resp = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=timeout)
                if resp.status_code == 200:
                    data = resp.json()
                    text = data["candidates"][0]["content"]["parts"][0]["text"]
                    return text
                elif resp.status_code in [503, 429]:
                    logger.warning(f"Model {model_name} HTTP {resp.status_code} verdi (deneme {attempt+1}), 1 sn bekleniyor.")
                    time.sleep(1.2)
                    last_error = f"HTTP {resp.status_code}: {resp.text[:120]}"
                    continue
                elif resp.status_code == 404:
                    last_error = f"HTTP 404: {resp.text[:120]}"
                    break
                else:
                    resp.raise_for_status()
            except requests.exceptions.RequestException as e:
                logger.warning(f"Model {model_name} çağrısında hata: {e}")
                last_error = str(e)
                time.sleep(0.5)
                continue

    raise RuntimeError(f"Tüm Gemini modelleri denendi fakat yanıt alınamadı. Son hata: {last_error}")

def extract_file_content(filename: str, file_bytes: bytes) -> tuple:
    """
    Yüklenen dosyanın içeriğini türüne göre çıkarır.
    Döndürür: (metin_icerik: str, inline_image_dict: Optional[dict])
    """
    ext = os.path.splitext(filename)[1].lower()

    # Görsel / Fotoğraf
    if ext in [".jpg", ".jpeg", ".png", ".webp"]:
        import base64
        mime = "image/jpeg" if ext in [".jpg", ".jpeg"] else ("image/png" if ext == ".png" else "image/webp")
        b64 = base64.b64encode(file_bytes).decode("utf-8")
        return f"[Görsel/Fotoğraf Dosyası: {filename}]", {"mime_type": mime, "data": b64}

    # Excel Tabloları (MEBBİS İSG Çıktıları)
    elif ext in [".xlsx", ".xls"]:
        import io
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        text_lines = []
        for sheet in wb.sheetnames[:3]:
            ws = wb[sheet]
            text_lines.append(f"--- Sayfa: {sheet} ---")
            for row in ws.iter_rows(max_row=120, values_only=True):
                non_empty = [str(cell) for cell in row if cell is not None]
                if non_empty:
                    text_lines.append(" | ".join(non_empty))
        return "\n".join(text_lines), None

    # PDF Rapor ve Tutanaklar
    elif ext == ".pdf":
        import io
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        pages_text = []
        for i, page in enumerate(reader.pages[:10]):
            txt = page.extract_text()
            if txt:
                pages_text.append(f"--- Sayfa {i+1} ---\n{txt}")
        return "\n".join(pages_text), None

    # Word Dokümanları
    elif ext == ".docx":
        import io
        import docx
        doc = docx.Document(io.BytesIO(file_bytes))
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        for t in doc.tables:
            for row in t.rows:
                row_txt = [c.text.strip() for c in row.cells if c.text.strip()]
                if row_txt:
                    paragraphs.append(" | ".join(row_txt))
        return "\n".join(paragraphs), None

    # Düz Metin / CSV
    else:
        try:
            return file_bytes.decode("utf-8"), None
        except UnicodeDecodeError:
            return file_bytes.decode("iso-8859-9", errors="replace"), None

def analyze_and_extract_mebbis_report(file_text: str, inline_image: dict, user_instruction: str, schools: list) -> dict:
    """
    Yüklenen MEBBİS raporunu, denetim tutanağını veya fotoğrafı analiz eder,
    ilgili okulu eşler ve veritabanı güncellemelerini JSON şemasıyla döker.
    """
    schools_ctx = ", ".join([f"{s['id']}:{s['kurum_adi']}" for s in schools])

    prompt = f"""
Sen Türkiye Cumhuriyeti MEB ve 6331 sayılı İSG Kanunu standartlarına hakim bir İSG ve MEBBİS Veri İşleme Başuzmanısın.
Muratpaşa İlçe MEM bünyesindeki 268 okulun MEBBİS İSG raporları, saha fotoğrafları ve periyodik muayene evrakları sana iletiliyor.

Kullanıcının Açıklaması / Talimatı:
\"\"\"{user_instruction or 'Belgedeki veya fotoğraftaki okul bilgilerini, periyodik kontrolleri ve güvenlik verilerini otomatik tespit et.'}\"\"\"

Belge Metni / OCR İçeriği:
\"\"\"{file_text[:8000]}\"\"\"

Sistemdeki Resmi Okul Listesi (Eşleştirme için referans al):
{schools_ctx}

GÖREVİN:
1. Yüklenen belge/fotoğraf ve kullanıcının açıklamasına göre hangi okul olduğunu tespit et ('matched_school_id' ve 'okul_adi_tespit').
2. Belgedeki periyodik kontrol tarihlerini (kazan, elektrik topraklama, paratoner, yangın tüpü, asansör, su deposu), muayene durumlarını ('Gecerli' veya 'Kusurlu') tespit et.
3. Varsa tatbikat tarihi, ilk yardımcı sayısı, güvenlik görevlisi, kamera sayısı veya öğrenci/öğretmen sayılarını tespit et.
4. 'periodic_controls_ekle' listesi içine eklenebilecek yeni periyodik kontrol muayenelerini yapılandır.
5. Aşağıdaki JSON şemasında eksiksiz ve geçerli bir JSON döndür. Tarihler YYYY-MM-DD formatında olmalıdır. Bulunmayan alanları null bırak.

JSON ŞEMASI:
{{
  "matched_school_id": integer veya null,
  "okul_adi_tespit": "Tespit edilen okul adı",
  "kurum_kodu_tespit": "Varsa kurum kodu" veya null,
  "belge_turu": "MEBBİS İSG Raporu" | "Periyodik Kontrol Belgesi" | "Saha Fotoğrafı / Tutanak" | "Fatura / Dolum Etiketi" | "Diğer",
  "guven_derecesi": "Yüksek" | "Orta" | "Düşük",
  "tespit_edilen_alanlar": {{
    "kazan_kontrol_tarihi": "YYYY-MM-DD" veya null,
    "kazan_durumu": "Gecerli" | "Kusurlu" veya null,
    "elektrik_topraklama_tarihi": "YYYY-MM-DD" veya null,
    "paratoner_kontrol_tarihi": "YYYY-MM-DD" veya null,
    "paratoner_durumu": "Gecerli" | "Kusurlu" veya null,
    "yangin_tupleri_kontrol_tarihi": "YYYY-MM-DD" veya null,
    "asansor_muayene_tarihi": "YYYY-MM-DD" veya null,
    "asansor_durumu": "Gecerli" | "Kusurlu" veya null,
    "su_deposu_temizlik_tarihi": "YYYY-MM-DD" veya null,
    "su_deposu_durumu": "Gecerli" | "Kusurlu" veya null,
    "adep_var_mi": "Var" | "Yok" veya null,
    "adep_tarihi": "YYYY-MM-DD" veya null,
    "tatbikat_tarihi": "YYYY-MM-DD" veya null,
    "tatbikat_turu": "Deprem Tahliye" | "Yangın" | "İkaz-Alarm" veya null,
    "tatbikat_katilimci_sayisi": integer veya null,
    "ilkyardimci_personel_sayisi": integer veya null,
    "ozel_guvenlik_sayisi": integer veya null,
    "toplam_kamera_sayisi": integer veya null,
    "ogrenci_sayisi": integer veya null,
    "ogretmen_sayisi": integer veya null,
    "mudur_ad_soyad": string veya null,
    "mudur_tel": string veya null
  }},
  "periodic_controls_ekle": [
    {{
      "ekipman_turu": "Kalorifer Kazanı" | "Yangın Söndürme Tüpleri" | "Elektrik İç Tesisat / Topraklama" | "Paratoner" | "Asansör" | "Su Deposu",
      "son_kontrol_tarihi": "YYYY-MM-DD",
      "gecerlilik_ayi": 12,
      "sonraki_kontrol_tarihi": "YYYY-MM-DD",
      "kontrol_yapan_kurum": "Yetkili Kuruluş / Firma",
      "rapor_sonucu": "Gecerli" | "Kusurlu",
      "notlar": "Muayene detayı"
    }}
  ],
  "ai_ozeti": "Kullanıcıya gösterilecek 1-2 cümlelik işlem özeti"
}}
"""
    parts = [{"text": prompt}]
    if inline_image:
        parts.append({"inline_data": inline_image})

    raw_json = call_gemini_multimodal(parts, json_mode=True)
    return json.loads(raw_json)


def analyze_incident_ai(raw_text: str, schools: list = None) -> dict:
    """
    Kullanıcının serbest metin olarak yazdığı ramak kala veya güvenlik olayını
    analiz eder, ilgili okulu, olay türünü, risk derecesini, mevzuat maddesini ve DÖF'ü çıkarır.
    """
    schools_context = ""
    if schools:
        # En fazla 60 okul ismi verelim ki context şişmesin
        school_names = [f"{s.get('id', '')}:{s.get('kurum_adi', '')}" for s in schools[:60]]
        schools_context = "Sistemdeki bazı okullar (ID:Ad formatında): " + ", ".join(school_names)

    prompt = f"""
Sen Türkiye Cumhuriyeti Milli Eğitim Bakanlığı ve 6331 sayılı İş Sağlığı ve Güvenliği Kanunu konusunda kıdemli bir İSG Başuzmanısın.
Muratpaşa İlçe MEM bünyesindeki okullarda meydana gelen kaza, ramak kala ve güvensiz durumları analiz ediyorsun.

Kullanıcıdan gelen serbest metin:
\"\"\"{raw_text}\"\"\"

{schools_context}

Görevin:
1. Bu metni analiz et.
2. Aşağıdaki JSON formatında eksiksiz ve geçerli bir JSON objesi döndür.
3. 'olay_turu' alanı yalnızca şu değerlerden biri olmalıdır:
   - "RAMAK_KALA"
   - "GUVENSIZ_DURUM"
   - "YARALANMALI_KAZA"
   - "UZUV_KAYBI_AGIR"
   - "YANGIN_PATLAMA"
   - "DIGER"
4. 'cde_seviye' alanı şu değerlerden biri olmalıdır:
   - "SEVIYE_1_OKUL" (Basit olaylar, ramak kala, hafif sıyrık)
   - "SEVIYE_2_ILCE" (Hastanelik vakalar, kırık, yangın başlangıcı, orta risk)
   - "SEVIYE_3_CDE" (Ağır kaza, uzuv kaybı, silahlı tehdit, tahliye)
5. 'termin_saat' aciliyetine göre 12, 24 veya 48 olmalıdır.
6. Eğer metinde okul adı geçiyorsa ve yukarıdaki listede eşleşen okul varsa 'matched_school_id' alanına okulun ID'sini yaz, yoksa null bırak.

JSON Şeması:
{{
  "matched_school_id": integer veya null,
  "olay_turu": "RAMAK_KALA" | "GUVENSIZ_DURUM" | "YARALANMALI_KAZA" | "UZUV_KAYBI_AGIR" | "YANGIN_PATLAMA" | "DIGER",
  "cde_seviye": "SEVIYE_1_OKUL" | "SEVIYE_2_ILCE" | "SEVIYE_3_CDE",
  "termin_saat": 12 | 24 | 48,
  "tehlike_kaynagi": "Kısa ve net tehlike tanımı (Örn: Hasarlı Bahçe Izgarası)",
  "olasilik": 1 ile 5 arası tam sayı,
  "siddet": 1 ile 5 arası tam sayı,
  "risk_skoru": olasilik * siddet,
  "risk_seviyesi": "Düşük" | "Orta" | "Yüksek" | "Çok Yüksek",
  "mevzuat_maddesi": "İlgili yönetmelik ve madde adı (Örn: İşyeri Bina ve Eklentileri Yönetmeliği Ek-1 Madde 7)",
  "onerilen_dof": "Alınması gereken düzeltici ve önleyici faaliyet (DÖF) önerisi",
  "resmi_aciklama": "Resmi kaza/olay tutanağına uygun, profesyonel dille yazılmış 1-2 cümlelik vaka açıklaması"
}}
"""
    raw_json = call_gemini(prompt, json_mode=True)
    return json.loads(raw_json)

def ask_isg_advisor_ai(question: str, school_context: str = "") -> str:
    """
    İSG Mevzuatı, MEB Okulum Temiz, ADEP, Risk Analizi ve Periyodik Kontroller
    konusunda uzman yapay zeka danışmanı.
    """
    system_prompt = f"""
Sen Muratpaşa İlçe Milli Eğitim Müdürlüğü'nün resmi İSG Yapay Zeka Asistanısın.
Uzmanlık Alanların:
- 6331 sayılı İş Sağlığı ve Güvenliği Kanunu
- MEB Okul ve Kurumlarında Acil Durumlar Hakkında Tebliğ & ADEP Hazırlama
- MEB & TSE Okulum Temiz Hijyen ve Enfeksiyon Önleme Kontrol Kılavuzu
- Binaların Yangından Korunması Hakkında Yönetmelik (Yangın tüpleri, kaçış yolları, tatbikatlar)
- İş Ekipmanlarının Kullanımında Sağlık ve Güvenlik Şartları Yönetmeliği (Kazan, asansör, elektrik, paratoner periyodik kontrolleri)
- Risk Değerlendirmesi Yönetmeliği (5x5 L Tipi ve Fine-Kinney Matrisi)

{school_context}

Kullanıcı Sorusu: \"\"\"{question}\"\"\"

Lütfen cevabını verirken:
1. Samimi, saygılı, bürokratik nezakete uygun ve net bir Türkçe kullan.
2. Mümkün olduğunda doğrudan ilgili Yönetmelik, Tebliğ veya Kanun maddesini referans göster.
3. Maddeler halinde pratik, uygulanabilir aksiyon adımları sun.
4. Yanıtın sonuna "Muratpaşa İlçe İSG Büro Koordinatörlüğü" adına iyi çalışmalar dileyen kısa bir not ekle.
"""
    return call_gemini(system_prompt, json_mode=False)

def analyze_feedback_request(
    gonderen_ad_soyad: str,
    unvan: str,
    okul_adi: str,
    talep_basligi: str,
    detayli_aciklama: str,
    ornek_yapi: str = ""
) -> dict:
    """
    Kullanıcıdan gelen portal geliştirme isteğini analiz ederek,
    Geliştirici Ramazan ERYILDIZ ve Antigravity IDE için teknik bir sprint task ve mimari prompt üretir.
    """
    prompt = f"""
Sen Muratpaşa İlçe MEM İSG Bilgi Yönetim Portalı Baş Yazılım Mimarı ve AI Eş-Geliştiricisisin.
Saha kullanıcısından bir geri bildirim ve yeni özellik talebi geldi:

GÖNDEREN: {gonderen_ad_soyad} ({unvan or 'Kullanıcı'})
KURUM / OKUL: {okul_adi or 'Genel İlçe Kullanıcısı'}
TALEP BAŞLIĞI: {talep_basligi}
DETAYLI AÇIKLAMA:
\"\"\"{detayli_aciklama}\"\"\"

EK ÖRNEK FORM / ŞABLON NOTU:
\"\"\"{ornek_yapi}\"\"\"

Lütfen bu talebi analiz et ve aşağıdaki JSON şemasına uygun olarak yanıt ver:

JSON Şeması:
{{
  "ai_ozeti": "Talebin özeti ve sisteme katacağı kurumsal değer (2-3 cümle)",
  "oneri_kategori": "YENI_DOKUMAN_SABLONU" | "YENI_VERI_GIRISI_FORMU" | "RAPOR_VE_ICMAL_TALEBI" | "ARAYUZ_KULLANILABILIRLIK" | "MEVZUAT_GUNCELLEME" | "HATA_BILDIRIMI" | "DIGER",
  "oncelik": "ACIL" | "YUKSEK" | "NORMAL" | "DUSUK",
  "veritabani_degisikligi_var_mi": true | false,
  "tahmini_mimari_etki": "Hangi tablolar veya endpoint'ler etkilenir (örn: schools tablosuna alan, yeni template vb.)",
  "teknik_sprint_promptu": "Antigravity IDE kodlama ajanına doğrudan kopyalanıp verilebilecek, net teknik talimat içeren prompt metni."
}}
"""
    try:
        raw_json = call_gemini(prompt, json_mode=True)
        return json.loads(raw_json)
    except Exception as e:
        logger.error(f"Feedback AI analizi hatası: {e}")
        return {
            "ai_ozeti": f"Kullanıcı talebi kaydedildi: {talep_basligi}",
            "oneri_kategori": "DIGER",
            "oncelik": "NORMAL",
            "veritabani_degisikligi_var_mi": False,
            "tahmini_mimari_etki": "Belirtilmedi",
            "teknik_sprint_promptu": f"Kullanıcı '{talep_basligi}' başlıklı bir geliştirme talebinde bulundu: {detayli_aciklama}"
        }

def bridge_mebbis_risk(
    bolum_unite: str,
    meslek_alani_veya_atolye: str,
    tehlike_kaynagi: str,
    risk_senaryosu: str,
    mevcut_onlemler: str,
    dof_onlem: str,
    risk_skoru: float,
    risk_seviyesi: str
) -> dict:
    """
    Sahada tespit edilen teknik riski MEBBİS İSG standart kütüphanesine eşleştirir,
    MEBBİS'e doğrudan kopyalanacak resmî teknik DÖF metnini ve 2024/38 ödenek kapsamını üretir.
    """
    prompt = f"""
Sen 6331 sayılı İSG Kanunu ve Millî Eğitim Bakanlığı MEBBİS İSG Modülü konusunda uzman A Sınıfı İş Güvenliği Uzmanısın.
Okul sahasında aşağıdaki risk ve tehlike tespit edildi:

BÖLÜM / ÜNİTE: {bolum_unite}
ATÖLYE / MESLEK ALANI: {meslek_alani_veya_atolye or 'Genel Bina/Eklenti'}
TEHLİKE KAYNAĞI: {tehlike_kaynagi}
OLASI TEHLİKE / RİSK SENARYOSU: {risk_senaryosu}
MEVCUT ÖNLEMLER: {mevcut_onlemler}
ALINMASI GEREKEN DÖF ÖNERİSİ: {dof_onlem}
HESAPLANAN RİSK SKORU: {risk_skoru} ({risk_seviyesi})

GÖREVİN:
1. MEBBİS İSG Kütüphanesi standart kodu (Örn: İSG-ELK-04, İSG-YNG-02, İSG-MAK-07, İSG-BINA-09 vb.) öner.
2. MEBBİS DÖF Açıklama Metni: Okul idaresi veya İSG uzmanı tarafından MEBBİS'e doğrudan kopyalanıp yapıştırılacak, mevzuat diliyle yazılmış, resmi ve profesyonel DÖF gerekçe metni hazırla.
3. İlgili yasal dayanak yönetmelik ve maddesini belirt.
4. Bu iş MEB 2024/38 sayılı Genelge kapsamında Bakanlıktan genel bütçe ödeneği talep edilecek nitelikte ağır fiziki/tesisat işi midir (true/false)?

JSON Şeması:
{{
  "mebbis_kutuphane_kodu": "STANDART-KOD-01",
  "mebbis_dof_aciklama_metni": "Resmi MEBBİS DÖF metni...",
  "dayanak_mevzuat": "İlgili yönetmelik ve madde",
  "odenek_talep_kapsaminda_mi": true | false,
  "ai_uzman_tavsiyesi": "Sahadaki uygulayıcıya 1 cümlelik pratik uzman tavsiyesi"
}}
"""
    try:
        raw_json = call_gemini(prompt, json_mode=True)
        return json.loads(raw_json)
    except Exception as e:
        logger.error(f"MEBBİS Bridge AI hatası: {e}")
        return {
            "mebbis_kutuphane_kodu": "İSG-GENEL-01",
            "mebbis_dof_aciklama_metni": f"{tehlike_kaynagi} kaynaklı risk için: {dof_onlem}",
            "dayanak_mevzuat": "6331 sayılı İş Sağlığı ve Güvenliği Kanunu Madde 4 ve 10",
            "odenek_talep_kapsaminda_mi": False,
            "ai_uzman_tavsiyesi": "Tespit edilen uygunsuzluk için ivedilikle DÖF başlatılmalıdır."
        }

def process_field_voice_and_vision_hazard(
    voice_transcript: str,
    image_bytes: bytes = None,
    image_mime: str = "image/jpeg",
    schools_cache: list = None,
    locked_school_id: int = None
) -> dict:
    """
    Sahada sesli not bırakan ve (varsa) fotoğraf çeken İSG uzmanının tespitini
    Gemini Multimodal yapay zekası ile analiz eder:
    - 268 okul arasından ilgili okulu tespit eder veya kilitli okulu baz alır,
    - En yakın aday okulları (candidate_schools) üretir,
    - Bölüm / Atölyeyi belirler,
    - Tehlike kaynağını ve olası risk senaryosunu oluşturur,
    - Fine-Kinney matris parametrelerini (O, F, Ş, Skor) hesaplar,
    - Standart MEBBİS İSG kütüphane kodunu ve resmî DÖF metnini üretir,
    - Yasal termin ve sorumluları belirler.
    """
    import base64
    import difflib
    from datetime import datetime, timedelta

    matched_school = None
    candidate_schools = []

    # 1. Uzman sayfada önceden okulu kilitlediyse doğrudan onu kullan
    if locked_school_id and schools_cache:
        try:
            l_id = int(locked_school_id)
            for s in schools_cache:
                if s.get("id") == l_id:
                    matched_school = s
                    candidate_schools = [{
                        "id": s["id"],
                        "kurum_kodu": s["kurum_kodu"],
                        "kurum_adi": s["kurum_adi"],
                        "score": 100.0
                    }]
                    break
        except (ValueError, TypeError):
            pass

    # 2. Kilitli okul yoksa gelişmiş fonetik & Levenshtein tabanlı okul sıralama motoru
    if not matched_school and schools_cache:
        UNIT_KEYWORDS = {'mutfak', 'atölye', 'atölyesi', 'atölyesindeki', 'laboratuvar', 'laboratuvarı', 'kazan', 'dairesi', 'yangın', 'dolabı', 'merdiveni', 'elektrik', 'panosu', 'koridor', 'bahçe', 'çatı', 'tuvalet', 'lavabo', 'bodrum', 'yemekhane', 'kantin', 'pansiyon', 'duvarda', 'fayans', 'fayanslar', 'dökülmüş', 'kırık', 'arızalı', 'kapı', 'pencere'}
        STOP_WORDS = {'ve', 'ile', 'ortaokulu', 'ilkokulu', 'lisesi', 'lisesindeyim', 'anadolu', 'mesleki', 'meslek', 'teknik', 'okulu', 'kurumu', 'merkezi', 'anaokulu', 'özel', 'şu', 'anda', 'buradayım', 'şimdi', 'okulundayım', 'okulunda'}
        PHONETIC_REPLACEMENTS = {
            'test tur': 'tez tur',
            'testur': 'tez tur',
            'teztur': 'tez tur',
            'falez': 'falez turizm',
            'serkan colkesen': 'serkan çölkesen',
            'cölkesen': 'çölkesen',
            'colkesen': 'çölkesen',
            'baro': 'antalya barosu',
            'metin nuran': 'metin nuran çakallıklı'
        }

        t_clean = voice_transcript.lower().replace("-", " ")
        for bad_p, good_p in PHONETIC_REPLACEMENTS.items():
            if bad_p in t_clean:
                t_clean = t_clean.replace(bad_p, good_p)

        words = [w for w in t_clean.split() if w not in UNIT_KEYWORDS and w not in STOP_WORDS and len(w) > 2]
        brand_phrase = ' '.join(words)

        scored = []
        for s in schools_cache:
            name_clean = s.get("kurum_adi", "").lower().replace("-", " ")
            name_words = [w for w in name_clean.split() if w not in STOP_WORDS and len(w) > 2]
            name_distinct = ' '.join(name_words)

            sim_phrase = difflib.SequenceMatcher(None, brand_phrase, name_distinct).ratio() if (brand_phrase and name_distinct) else 0.0

            token_score = 0.0
            for w in words:
                for nw in name_words:
                    r = difflib.SequenceMatcher(None, w, nw).ratio()
                    if r >= 0.8:
                        token_score += 30.0 * r
                    elif r >= 0.6:
                        token_score += 10.0 * r

            total = (sim_phrase * 50.0) + token_score
            scored.append((s, total))

        scored.sort(key=lambda x: x[1], reverse=True)
        top_candidates = [item for item in scored[:4] if item[1] >= 18.0]
        if top_candidates:
            matched_school = top_candidates[0][0]
            candidate_schools = [{
                "id": s["id"],
                "kurum_kodu": s["kurum_kodu"],
                "kurum_adi": s["kurum_adi"],
                "score": round(sc, 1)
            } for s, sc in top_candidates]

    prompt = f"""
Sen Antalya Muratpaşa İlçe Millî Eğitim Müdürlüğü A Sınıfı Baş İSG Denetçisi ve MEBBİS İSG Modülü uzmanısın.
Saha denetimindeki İSG uzmanı cep telefonundan konuşarak sesli not bıraktı ve olay yerinin fotoğrafını çekti (veya sisteme iletti).

UZMANIN SESLİ SAHA İFADESİ:
"{voice_transcript}"

GÖREVİN:
1. Okul Tespiti: Sesli notta geçen okul adını veya kısaltmasını tespit et. (Örn: 'teztur meslek' -> 'Tez-Tur Turizm Mesleki ve Teknik Anadolu Lisesi', 'gazi' -> 'Gazi Anadolu Lisesi', 'serkan çölkesen' -> 'Şehit Serkan Çölkesen Ticaret MTAL').
2. Bölüm / Atölye: Hangi birimde veya atölyede olduğunu belirle (Örn: Mutfak Atölyesi, Kazan Dairesi, Elektrik Panosu, Kimya Laboratuvarı, Yangın Merdiveni, Spor Salonu).
3. Tehlike Kaynağı: Kısa, somut tehlike tanımı (Örn: Duvarda dökülen ve kırılan fayans kaplamalar).
4. Olası Tehlike ve Risk Senaryosu: Durumun doğuracağı tehlike ve olası kaza/sağlık riski (Hijyen, kesilme, kayma, elektrik çarpması vb.).
5. Fine-Kinney Parametreleri:
   - Olasılık (O): 0.2 (Hemen hemen imkansız), 0.5 (Çok nadir), 1.0 (Oldukça az), 3.0 (Nadir fakat olabilir), 6.0 (Kuvvetle muhtemel), 10.0 (Beklenir/Kesin)
   - Frekans (F): 0.5 (Yılda bir), 1.0 (Yılda birkaç kez), 2.0 (Ayda bir), 3.0 (Haftada bir), 6.0 (Günde bir), 10.0 (Sürekli)
   - Şiddet (Ş): 1 (İlk yardım), 3 (Ayakta tedavi/pansuman), 7 (Hafif yaralanma), 15 (Ciddi yaralanma/kırık), 40 (Kalıcı sakatlık/tek ölüm), 100 (Birden fazla ölüm)
   - Skor = O * F * Ş
6. Standart MEBBİS İSG Kütüphane Kodu (Örn: İSG-MUT-04, İSG-BINA-12, İSG-ELK-03, İSG-MAK-06, İSG-LAB-01 vb.).
7. MEBBİS DÖF Açıklama Metni: Okul idaresi veya uzmanın MEBBİS İSG modülüne doğrudan kopyalayıp yapıştıracağı, 6331 sayılı Kanun ve ilgili yönetmelik diline uygun, resmi, emir kipiyle yazılmış DÖF önlem metni.
8. Yasal Dayanak Mevzuat (Örn: İş Ekipmanlarının Kullanımında Sağlık ve Güvenlik Şartları Yönetmeliği, Binaların Yangından Korunması Hakkında Yönetmelik vb.).
9. Termin Süresi (Gün cinsinden: 7, 15, 30 veya 60).
10. Sorumlular (Örn: Okul Yönetimi / Atölye Şefi / Nöbetçi Müdür Yardımcısı).
11. MEB 2024/38 Sayılı Ödenek Kapsamında mı (Bakanlıktan onarım bütçesi istenecek yapısal bir iş mi? true/false).
12. Görsel Tespit Notu: (Eğer fotoğraf yüklendiyse fotoğraftaki durumu kısaca özetle, fotoğraf yoksa 'Sesli beyan üzerinden analiz yapıldı' de).

YANITINI SADECE AŞAĞIDAKİ JSON ŞEMASINDA DÖNDÜR:
{{
  "school_name_detected": "Okulun tespit edilen adı",
  "unit_or_workshop": "Bölüm veya Atölye",
  "bolum_unite_kategori": "ATOLYE_LABORATUVAR | KAZAN_DAIRESI | BINA_GENEL | ELEKTRIK_TRAFO | YANGIN_YOLU",
  "tehlike_kaynagi": "Kısa tehlike kaynağı",
  "risk_senaryosu": "Olası risk senaryosu ve doğuracağı zarar",
  "fine_kinney": {{
    "olasilik": 6.0,
    "frekans": 6.0,
    "siddet": 3.0,
    "skor": 108.0,
    "seviye": "Önemli Risk (70 - 200)",
    "seviye_kod": "YUKSEK_ONEMLI"
  }},
  "mebbis_kutuphane_kodu": "İSG-MUT-04",
  "mebbis_dof_aciklama_metni": "Resmî MEBBİS DÖF metni...",
  "dayanak_mevzuat": "İlgili yönetmelik adı",
  "termin_gun": 15,
  "sorumlular": "Okul Yönetimi & Atölye Şefi",
  "odenek_talep_kapsaminda_mi": true,
  "gorsel_analiz_notu": "Fotoğrafta görülen hasar/uygunsuzluk özeti"
}}
"""

    parts = [{"text": prompt}]
    if image_bytes:
        try:
            b64_img = base64.b64encode(image_bytes).decode("utf-8")
            parts.append({
                "inline_data": {
                    "mime_type": image_mime or "image/jpeg",
                    "data": b64_img
                }
            })
        except Exception as img_err:
            logger.warning(f"Görsel base64 çevrim hatası: {img_err}")

    today = datetime.now()
    default_termin = (today + timedelta(days=15)).strftime("%Y-%m-%d")

    try:
        raw_json = call_gemini_multimodal(parts, json_mode=True, timeout=30)
        data = json.loads(raw_json)
        
        # Termin tarihi hesapla
        days = int(data.get("termin_gun") or 15)
        data["termin_tarihi"] = (today + timedelta(days=days)).strftime("%Y-%m-%d")
        
        # Okul eşleşmesini ekle
        if matched_school:
            data["matched_school"] = {
                "id": matched_school.get("id"),
                "kurum_kodu": matched_school.get("kurum_kodu"),
                "kurum_adi": matched_school.get("kurum_adi")
            }
        else:
            data["matched_school"] = None

        data["candidate_schools"] = candidate_schools

        return {"success": True, "data": data}
    except Exception as e:
        logger.error(f"Saha Sesli & Görsel AI hatası: {e}")
        # Akıllı yedek (Fallback)
        fallback_data = {
            "school_name_detected": matched_school.get("kurum_adi") if matched_school else "Belirtilmedi",
            "unit_or_workshop": "Saha Tespit Birimi",
            "bolum_unite_kategori": "ATOLYE_LABORATUVAR",
            "tehlike_kaynagi": voice_transcript[:80],
            "risk_senaryosu": f"Sözlü saha tespiti: {voice_transcript}",
            "fine_kinney": {
                "olasilik": 6.0,
                "frekans": 6.0,
                "siddet": 3.0,
                "skor": 108.0,
                "seviye": "Önemli Risk / DÖF Başlatılmalı (70 - 200)",
                "seviye_kod": "YUKSEK_ONEMLI"
            },
            "mebbis_kutuphane_kodu": "İSG-SAHA-01",
            "mebbis_dof_aciklama_metni": f"Sahada tespit edilen uygunsuzluk ({voice_transcript}) giderilmeli ve koruyucu önlemler alınmalıdır.",
            "dayanak_mevzuat": "6331 sayılı İş Sağlığı ve Güvenliği Kanunu Madde 4, 10 ve 16",
            "termin_gun": 15,
            "termin_tarihi": default_termin,
            "sorumlular": "Okul Yönetimi / İSG Kurulu",
            "odenek_talep_kapsaminda_mi": False,
            "gorsel_analiz_notu": "Sesli beyan üzerinden kural tabanlı analiz yapıldı.",
            "matched_school": {
                "id": matched_school.get("id"),
                "kurum_kodu": matched_school.get("kurum_kodu"),
                "kurum_adi": matched_school.get("kurum_adi")
            } if matched_school else None
        }
        return {"success": True, "data": fallback_data, "fallback": True}



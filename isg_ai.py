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

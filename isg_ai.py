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
    "gemini-3.8-flash",
    "gemini-flash-latest",
    "gemini-2.5-flash"
]

def call_gemini(prompt: str, json_mode: bool = False, timeout: int = 25) -> str:
    """Gemini API'sini çağırır. Model meşgulse (503/404) otomatik olarak yedek modele geçer."""
    api_key = get_api_key()
    if not api_key:
        raise ValueError("GEMINI_API_KEY bulunamadı.")

    payload = {
        "contents": [{"parts": [{"text": prompt}]}]
    }
    if json_mode:
        payload["generationConfig"] = {"response_mime_type": "application/json"}

    last_error = None
    for model_name in CANDIDATE_MODELS:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        try:
            resp = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=timeout)
            if resp.status_code == 200:
                data = resp.json()
                text = data["candidates"][0]["content"]["parts"][0]["text"]
                return text
            elif resp.status_code in [404, 503, 429]:
                logger.warning(f"Model {model_name} HTTP {resp.status_code} verdi, bir sonraki modele geçiliyor.")
                last_error = f"HTTP {resp.status_code}: {resp.text[:120]}"
                continue
            else:
                resp.raise_for_status()
        except requests.exceptions.RequestException as e:
            logger.warning(f"Model {model_name} çağrısında hata: {e}")
            last_error = str(e)
            continue

    raise RuntimeError(f"Tüm Gemini modelleri denendi fakat yanıt alınamadı. Son hata: {last_error}")

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

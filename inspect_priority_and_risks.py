import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

ISG_DIR = r"C:\iş_güvenliği"

def inspect_oncelikli_muratpasa():
    print("=" * 65)
    print(" 1. ÖNCELİKLİ OKULLAR LİSTESİ (MURATPAŞA FİLTRESİ)")
    print("=" * 65)
    fp = os.path.join(ISG_DIR, "ÖNCELİKLİ OKULLAR LİSTESİ ANTALYA İL GENELİ2026.xlsx")
    if not os.path.exists(fp):
        print("Dosya bulunamadı.")
        return

    wb = openpyxl.load_workbook(fp, data_only=True)
    ws = wb['İL GENELİ']
    
    muratpasa_schools = []
    # Sütunlar: SIRA, IL, ILCE, KURUM_KODU, KURUM_ADI, KADEME, KOLLUK_SORUMLULUK_ALANI, MEB_OLCEK_UANI, KOLLUK UANI 2, TOPLAM_PUAN, DERECE
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[2]:
            continue
        ilce = str(row[2]).strip().upper()
        if "MURATPAŞA" in ilce or "MURATPASA" in ilce:
            muratpasa_schools.append({
                "kod": row[3],
                "ad": row[4],
                "kademe": row[5],
                "kolluk": row[6],
                "meb_puan": row[7],
                "kolluk_puan": row[8],
                "toplam_puan": row[9],
                "derece": row[10]
            })

    print(f"Toplam Muratpaşa Kurumu: {len(muratpasa_schools)}")
    for s in muratpasa_schools[:15]:
        print(f"  * {s['kod']} - {s['ad']} | Derece: {s['derece']}. Derece | Toplam Puan: {s['toplam_puan']} (Kolluk: {s['kolluk']})")
    
    return muratpasa_schools

def inspect_risk_reports():
    print("\n" + "=" * 65)
    print(" 2. C:\\iş_güvenliği\\risk_raporları KLASÖRÜNDEKİ 6 DOSYA")
    print("=" * 65)
    rr_dir = os.path.join(ISG_DIR, "risk_raporları")
    if not os.path.exists(rr_dir):
        return
    for f in os.listdir(rr_dir):
        fp = os.path.join(rr_dir, f)
        print(f"\nDosya: {f} ({os.path.getsize(fp)} bayt)")
        # İlk 1000 karakteri oku (html veya csv veya text)
        try:
            with open(fp, "r", encoding="utf-8", errors="ignore") as tf:
                content = tf.read(500)
                # Kurum adı bulmaya çalış
                lines = [l.strip() for l in content.splitlines() if l.strip()]
                for l in lines[:5]:
                    print(f"  | {l[:100]}")
        except Exception as e:
            print(f"  Okuma hatası: {e}")

if __name__ == "__main__":
    inspect_oncelikli_muratpasa()
    inspect_risk_reports()

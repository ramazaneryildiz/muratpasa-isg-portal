import os
import sys
from auth.credential_store import get_credential_store
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding='utf-8')

def main():
    print("=" * 65)
    print(" 1. C:\\iş_güvenliği\\yanıtlar KLASÖRÜ İNCELENİYOR")
    print("=" * 65)
    yanitlar_dir = r"C:\iş_güvenliği\yanıtlar"
    if os.path.exists(yanitlar_dir):
        files = os.listdir(yanitlar_dir)
        print(f"Toplam dosya: {len(files)}")
        for f in files:
            fp = os.path.join(yanitlar_dir, f)
            print(f"  * {f} ({os.path.getsize(fp)} bayt)")

    print("\n" + "=" * 65)
    print(" 2. VCF REHBERİ TÜM KİŞİLER İNCELENİYOR")
    print("=" * 65)
    store = get_credential_store()
    creds = store.get_credential('ramazaneryildiz@gmail.com')
    drive = build('drive', 'v3', credentials=creds)

    res = drive.files().list(
        q="name contains 'Sorumlu_Okullar_Rehber' and trashed = false",
        fields="files(id, name)"
    ).execute()
    files = res.get('files', [])
    if files:
        f = files[0]
        content = drive.files().get_media(fileId=f['id']).execute()
        text = content.decode('utf-8', errors='ignore')
        cards = text.split("BEGIN:VCARD")
        print(f"Toplam VCard Kişi Sayısı: {len(cards)-1}")
        for i, card in enumerate(cards[1:], 1):
            lines = card.strip().splitlines()
            fn = next((l.replace("FN:", "").strip() for l in lines if l.startswith("FN:")), "")
            org = next((l.replace("ORG:", "").strip() for l in lines if l.startswith("ORG:")), "")
            tel = next((l.replace("TEL;TYPE=CELL,VOICE:", "").strip() for l in lines if "TEL" in l), "")
            title = next((l.replace("TITLE:", "").strip() for l in lines if l.startswith("TITLE:")), "")
            print(f"  [{i}] {fn} | {title} | {tel} | {org}")

    print("\n" + "=" * 65)
    print(" 3. GOOGLE FORM YANITLAR E-TABLOSU (1-xYvDp4oaV5TciGeSzizJhF76Sh3bsrtRaVCfSGzT-c)")
    print("=" * 65)
    try:
        # Spreadsheet sayfalarını kontrol et
        form_sheet_id = "1-xYvDp4oaV5TciGeSzizJhF76Sh3bsrtRaVCfSGzT-c"
        content = drive.files().export(fileId=form_sheet_id, mimeType='text/csv').execute()
        csv_text = content.decode('utf-8', errors='ignore')
        lines = csv_text.splitlines()
        print(f"Form E-Tablosu Satır Sayısı: {len(lines)}")
        for l in lines[:10]:
            print(f"  | {l[:100]}")
    except Exception as e:
        print(f"Hata: {e}")

if __name__ == "__main__":
    main()

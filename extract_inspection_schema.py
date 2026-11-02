import os
import sys
import io
import json
from auth.credential_store import get_credential_store
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding='utf-8')

DATA_DIR = r"c:\isg\data\inspection_reference"
os.makedirs(DATA_DIR, exist_ok=True)

def main():
    store = get_credential_store()
    creds = store.get_credential('ramazaneryildiz@gmail.com')
    drive = build('drive', 'v3', credentials=creds)

    print("=" * 65)
    print(" MÜFETTİŞ DENETİM DOSYASI & VERİ TOPLAMA REFERANSLARI İNCELENİYOR")
    print("=" * 65)

    folder_id = "1NtlMlzlugxYmNj5eOenR3dt6-audsf4S"
    res = drive.files().list(
        q=f"'{folder_id}' in parents and trashed = false",
        fields="files(id, name, mimeType, size)"
    ).execute()

    files = res.get('files', [])
    print(f"Klasörde {len(files)} dosya tespit edildi:\n")

    for f in files:
        name = f['name']
        fid = f['id']
        mime = f['mimeType']
        print(f"[*] İnceleniyor: {name} ({mime})")

        # Google Docs ise text olarak export et
        if mime == 'application/vnd.google-apps.document':
            try:
                content = drive.files().export(fileId=fid, mimeType='text/plain').execute()
                text = content.decode('utf-8', errors='ignore')
                out_path = os.path.join(DATA_DIR, f"{name}.txt")
                with open(out_path, "w", encoding="utf-8") as out_f:
                    out_f.write(text)
                print(f"    -> TXT olarak kaydedildi ({len(text)} karakter). Özet:")
                lines = [l.strip() for l in text.splitlines() if l.strip()]
                for l in lines[:15]:
                    print(f"       | {l[:90]}")
                print()
            except Exception as e:
                print(f"    -> Export hatası: {e}")

        # Google Sheets ise csv olarak export et
        elif mime == 'application/vnd.google-apps.spreadsheet':
            try:
                content = drive.files().export(fileId=fid, mimeType='text/csv').execute()
                text = content.decode('utf-8', errors='ignore')
                out_path = os.path.join(DATA_DIR, f"{name}.csv")
                with open(out_path, "w", encoding="utf-8") as out_f:
                    out_f.write(text)
                print(f"    -> CSV olarak kaydedildi ({len(text)} karakter).")
                lines = [l.strip() for l in text.splitlines() if l.strip()]
                for l in lines[:8]:
                    print(f"       | {l[:90]}")
                print()
            except Exception as e:
                print(f"    -> Export hatası: {e}")

        # Binary (docx, xlsx) ise doğrudan indir
        elif 'openxmlformats' in mime:
            try:
                content = drive.files().get_media(fileId=fid).execute()
                ext = ".xlsx" if "sheet" in mime else ".docx"
                out_path = os.path.join(DATA_DIR, f"{name}")
                if not out_path.endswith(ext):
                    out_path += ext
                with open(out_path, "wb") as out_f:
                    out_f.write(content)
                print(f"    -> İndirildi: {out_path} ({len(content)} bayt)\n")
            except Exception as e:
                print(f"    -> İndirme hatası: {e}")

if __name__ == "__main__":
    main()

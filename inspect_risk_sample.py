import sys
import io
import csv
from auth.credential_store import get_credential_store
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding='utf-8')

def main():
    store = get_credential_store()
    creds = store.get_credential('ramazaneryildiz@gmail.com')
    drive = build('drive', 'v3', credentials=creds)

    res = drive.files().list(
        q="name contains 'Risk_Analiz_Raporu_764955' and trashed = false",
        fields="files(id, name)"
    ).execute()
    files = res.get('files', [])
    if files:
        f = files[0]
        print(f"\n[+] Risk Tablosu İndiriliyor: {f['name']} (ID: {f['id']})")
        content = drive.files().export(
            fileId=f['id'],
            mimeType='text/csv'
        ).execute()

        csv_text = content.decode('utf-8', errors='ignore')
        reader = csv.reader(io.StringIO(csv_text))
        rows = list(reader)
        print(f"Toplam Satır: {len(rows)}")
        for i, row in enumerate(rows[:15]):
            print(f"  [{i+1}] {row}")

if __name__ == "__main__":
    main()

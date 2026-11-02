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

    print("=" * 65)
    print(" GOOGLE DRIVE API İLE TABLO VERİSİ ÇEKİLİYOR")
    print("=" * 65)

    # 1. 283 Kurum Tablosu
    res = drive.files().list(
        q="name contains '283_Kurum' and trashed = false",
        fields="files(id, name)"
    ).execute()
    files = res.get('files', [])
    if files:
        f = files[0]
        print(f"\n[+] Tablo İndiriliyor: {f['name']} (ID: {f['id']})")
        content = drive.files().export(
            fileId=f['id'],
            mimeType='text/csv'
        ).execute()

        # CSV olarak oku
        csv_text = content.decode('utf-8', errors='ignore')
        reader = csv.reader(io.StringIO(csv_text))
        rows = list(reader)
        print(f"Toplam Satır: {len(rows)}")
        for i, row in enumerate(rows[:10]):
            print(f"  [{i+1}] {row}")

    # 2. Sorumlu Okullar İletişim Tablosu
    res2 = drive.files().list(
        q="name contains 'Sorumlu_Okullar_Isveren' and trashed = false",
        fields="files(id, name)"
    ).execute()
    files2 = res2.get('files', [])
    if files2:
        f2 = files2[0]
        print(f"\n[+] İletişim Tablosu İndiriliyor: {f2['name']} (ID: {f2['id']})")
        content2 = drive.files().export(
            fileId=f2['id'],
            mimeType='text/csv'
        ).execute()
        csv_text2 = content2.decode('utf-8', errors='ignore')
        reader2 = csv.reader(io.StringIO(csv_text2))
        rows2 = list(reader2)
        print(f"Toplam Satır: {len(rows2)}")
        for i, row in enumerate(rows2[:10]):
            print(f"  [{i+1}] {row}")

if __name__ == "__main__":
    main()

import sys
import json
from auth.credential_store import get_credential_store
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding='utf-8')

def main():
    store = get_credential_store()
    creds = store.get_credential('ramazaneryildiz@gmail.com')
    sheets = build('sheets', 'v4', credentials=creds)
    drive = build('drive', 'v3', credentials=creds)

    print("=" * 65)
    print(" MURATPAŞA OKUL VE İLETİŞİM TABLOLARI İNCELEMESİ")
    print("=" * 65)

    # 1. 283 Kurum Listesi dosyasının ID'sini bulalım
    res = drive.files().list(
        q="name contains '283_Kurum' and trashed = false",
        fields="files(id, name)"
    ).execute()
    files = res.get('files', [])
    if files:
        file_id = files[0]['id']
        file_name = files[0]['name']
        print(f"\n[+] Tablo Bulundu: {file_name} (ID: {file_id})")
        # İlk sayfanın ilk 10 satırını oku
        sheet_data = sheets.spreadsheets().values().get(
            spreadsheetId=file_id,
            range="A1:Z10"
        ).execute()
        rows = sheet_data.get('values', [])
        print(f"  Toplam okunan örnek satır: {len(rows)}")
        for i, row in enumerate(rows):
            print(f"  Satır {i+1}: {row}")

    # 2. Sorumlu Okullar İletişim Listesi dosyasının ID'sini bulalım
    res2 = drive.files().list(
        q="name contains 'Sorumlu_Okullar_Isveren' and trashed = false",
        fields="files(id, name)"
    ).execute()
    files2 = res2.get('files', [])
    if files2:
        file_id2 = files2[0]['id']
        file_name2 = files2[0]['name']
        print(f"\n[+] İletişim Tablosu Bulundu: {file_name2} (ID: {file_id2})")
        sheet_data2 = sheets.spreadsheets().values().get(
            spreadsheetId=file_id2,
            range="A1:Z10"
        ).execute()
        rows2 = sheet_data2.get('values', [])
        print(f"  Toplam okunan örnek satır: {len(rows2)}")
        for i, row in enumerate(rows2):
            print(f"  Satır {i+1}: {row}")

if __name__ == "__main__":
    main()

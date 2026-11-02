import sys
import json
from auth.credential_store import get_credential_store
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding='utf-8')

def list_folder_contents(drive, folder_id, folder_name, indent=0):
    prefix = "  " * indent
    print(f"{prefix}📁 [{folder_name}] (ID: {folder_id})")
    
    # Alt klasörleri ve dosyaları çek
    q = f"'{folder_id}' in parents and trashed = false"
    res = drive.files().list(
        q=q,
        fields="files(id, name, mimeType, size, modifiedTime)",
        pageSize=100
    ).execute()
    
    items = res.get('files', [])
    for it in items:
        if it['mimeType'] == 'application/vnd.google-apps.folder':
            list_folder_contents(drive, it['id'], it['name'], indent + 1)
        else:
            size_kb = int(it.get('size', 0)) / 1024 if it.get('size') else 0
            print(f"{prefix}  📄 {it['name']} ({it['mimeType']}, {size_kb:.1f} KB)")

def main():
    store = get_credential_store()
    creds = store.get_credential('ramazaneryildiz@gmail.com')
    drive = build('drive', 'v3', credentials=creds)

    print("=" * 65)
    print(" İSG ANA KLASÖRÜ VE ALT DİZİNLERİ DETAYLI ANALİZİ")
    print("=" * 65)

    isg_main_id = "1Mwyd2JkO5Na4v8h7N_vhvJY1iX-w5bHC"
    list_folder_contents(drive, isg_main_id, "İSG Ana Klasörü")

    # Ayrıca diğer tespit edilen önemli İSG klasörlerini listeleyelim:
    other_folders = [
        ("Öncelikli Okullar", "1lo41FQOPmPJgPD7Gr-CGu2XsYYjuIOsd"),
        ("Okul Güvenliği Müfettiş Denetim Dosyası", "1NtlMlzlugxYmNj5eOenR3dt6-audsf4S"),
        ("Okul Risk Analizleri ve Saha Dosyaları", "1pybTbDzkzvgwJfyn4Q1ZtMyhnCEv7qZq"),
    ]
    for name, fid in other_folders:
        print("\n" + "-" * 50)
        list_folder_contents(drive, fid, name)

if __name__ == "__main__":
    main()

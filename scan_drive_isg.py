import sys
import json
from auth.credential_store import get_credential_store
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding='utf-8')

def main():
    store = get_credential_store()
    creds = store.get_credential('ramazaneryildiz@gmail.com')
    drive = build('drive', 'v3', credentials=creds)

    print("=" * 60)
    print(" GOOGLE DRIVE İSG KLASÖR VE DOSYA TARAMASI")
    print("=" * 60)

    # 1. Önce tüm klasörleri arayalım
    folder_query = "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    res_folders = drive.files().list(
        q=folder_query,
        fields="files(id, name, parents, modifiedTime)",
        pageSize=100
    ).execute()
    folders = res_folders.get('files', [])
    print(f"\n[+] Toplam bulunan klasör sayısı: {len(folders)}")
    for f in folders:
        print(f"  📁 Klasör: {f['name']} (ID: {f['id']})")

    # 2. İsimde İSG, okul, muratpasa, adep, risk vb. geçen tüm dosyaları arayalım
    print("\n[+] İSG / Okul / İlgili Dosyalar Taranıyor...")
    keywords = ["isg", "İSG", "okul", "Okul", "muratpaşa", "Muratpasa", "adep", "ADEP", "risk", "Risk", "acil", "Acil", "rehber", "Rehber"]
    query_parts = [f"name contains '{k}'" for k in keywords]
    file_query = f"({' or '.join(query_parts)}) and trashed = false"
    
    res_files = drive.files().list(
        q=file_query,
        fields="files(id, name, mimeType, size, modifiedTime, parents)",
        pageSize=100
    ).execute()
    files = res_files.get('files', [])
    print(f"[+] Eşleşen dosya sayısı: {len(files)}\n")
    for fl in files:
        size_kb = int(fl.get('size', 0)) / 1024 if fl.get('size') else 0
        print(f"  📄 Dosya: {fl['name']}")
        print(f"     Tür: {fl['mimeType']} | Boyut: {size_kb:.1f} KB | Tarih: {fl.get('modifiedTime')}")
        print(f"     ID: {fl['id']} | Parents: {fl.get('parents', [])}\n")

    # 3. Kök dizindeki genel dosyaların özeti (ilk 30)
    print("\n[+] Drive Genel Dosya Listesi (Son Güncellenenler):")
    res_all = drive.files().list(
        q="trashed = false",
        fields="files(id, name, mimeType, modifiedTime)",
        pageSize=30
    ).execute()
    for item in res_all.get('files', []):
        print(f"  * {item['name']} ({item['mimeType']})")

if __name__ == "__main__":
    main()

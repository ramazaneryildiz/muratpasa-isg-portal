import sys
from auth.credential_store import get_credential_store
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding='utf-8')

def main():
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
        print(f"\n[+] VCF İndiriliyor: {f['name']} (ID: {f['id']})")
        content = drive.files().get_media(fileId=f['id']).execute()
        text = content.decode('utf-8', errors='ignore')
        lines = text.splitlines()
        print(f"Toplam Satır: {len(lines)}")
        for l in lines[:25]:
            print(f"  {l}")

if __name__ == "__main__":
    main()

import sys
import json
from auth.credential_store import get_credential_store
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding='utf-8')

def parse_header(headers, name):
    for h in headers:
        if h['name'].lower() == name.lower():
            return h['value']
    return ""

def main():
    store = get_credential_store()
    creds = store.get_credential('ramazaneryildiz@gmail.com')
    gmail = build('gmail', 'v1', credentials=creds)

    res = gmail.users().messages().list(
        userId='me',
        q='is:unread in:inbox -in:trash -in:spam',
        maxResults=20
    ).execute()

    messages = res.get('messages', [])
    print(f"Toplam listelenen son okunmamış mail sayısı: {len(messages)}")

    results = []
    for m in messages:
        msg = gmail.users().messages().get(
            userId='me',
            id=m['id'],
            format='metadata',
            metadataHeaders=['From', 'Subject', 'Date']
        ).execute()

        headers = msg.get('payload', {}).get('headers', [])
        sender = parse_header(headers, 'From')
        subject = parse_header(headers, 'Subject')
        date = parse_header(headers, 'Date')
        snippet = msg.get('snippet', '')

        results.append({
            "id": m['id'],
            "from": sender,
            "subject": subject,
            "date": date,
            "snippet": snippet
        })

    for idx, r in enumerate(results, 1):
        print(f"\n[{idx}] Kimden: {r['from']}")
        print(f"    Konu: {r['subject']}")
        print(f"    Tarih: {r['date']}")
        print(f"    Özet/İçerik: {r['snippet']}")

if __name__ == "__main__":
    main()

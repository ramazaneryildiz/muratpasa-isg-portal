import sys
import json
from collections import Counter
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

    # 1. Sahibinden araması (hem genel hem spam/çöp kutusu dahil)
    print("=== 1. Sahibinden Taraması ===")
    s_queries = ['from:sahibinden.com', 'subject:sahibinden', 'sahibinden']
    for sq in s_queries:
        res = gmail.users().messages().list(userId='me', q=sq, maxResults=10).execute()
        msgs = res.get('messages', [])
        print(f"Sorgu '{sq}': {len(msgs)} sonuç")
        for m in msgs[:3]:
            try:
                msg = gmail.users().messages().get(userId='me', id=m['id'], format='metadata', metadataHeaders=['From', 'Subject', 'Date']).execute()
                headers = msg.get('payload', {}).get('headers', [])
                sender = parse_header(headers, 'From')
                subject = parse_header(headers, 'Subject')
                date = parse_header(headers, 'Date')
                print(f"   * Kimden: {sender} | Konu: {subject} | Tarih: {date}")
            except Exception:
                pass

    # 2. Tanıtım / Promosyonlar (İlk 150)
    print("\n=== 2. Tanıtım / Reklam Mailleri (Promotions - İlk 150) ===")
    p_res = gmail.users().messages().list(userId='me', q='category:promotions', maxResults=150).execute()
    p_msgs = p_res.get('messages', [])
    senders_p = Counter()
    samples_p = {}
    for m in p_msgs:
        try:
            msg = gmail.users().messages().get(userId='me', id=m['id'], format='metadata', metadataHeaders=['From', 'Subject']).execute()
            headers = msg.get('payload', {}).get('headers', [])
            sender = parse_header(headers, 'From')
            subject = parse_header(headers, 'Subject')
            senders_p[sender] += 1
            if sender not in samples_p:
                samples_p[sender] = []
            if len(samples_p[sender]) < 2:
                samples_p[sender].append(subject)
        except Exception:
            continue

    print(f"Toplam incelenen Tanıtım maili: {len(p_msgs)}")
    print("En Çok Gönderenler:")
    for sender, count in senders_p.most_common(12):
        print(f"  - [{count} adet] {sender}")
        sample_str = " | ".join(samples_p.get(sender, []))
        print(f"    Örnek: {sample_str[:80]}")

    # 3. Güncellemeler (Updates) ve Sosyal (Social)
    print("\n=== 3. Sosyal ve Güncellemeler Kategorileri ===")
    for cat in ['category:social', 'category:updates']:
        c_res = gmail.users().messages().list(userId='me', q=cat, maxResults=50).execute()
        c_msgs = c_res.get('messages', [])
        c_counter = Counter()
        for m in c_msgs:
            try:
                msg = gmail.users().messages().get(userId='me', id=m['id'], format='metadata', metadataHeaders=['From']).execute()
                headers = msg.get('payload', {}).get('headers', [])
                c_counter[parse_header(headers, 'From')] += 1
            except Exception:
                continue
        print(f"\n{cat} (Taranan: {len(c_msgs)}):")
        for sender, count in c_counter.most_common(5):
            print(f"  - [{count} adet] {sender}")

if __name__ == "__main__":
    main()

import sys
import time
from auth.credential_store import get_credential_store
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

sys.stdout.reconfigure(encoding='utf-8')

TARGET_CATEGORIES = {
    "Temu Reklamları": {
        "query": "(from:temu.com OR from:temuemail.com) -in:trash",
        "filter_from": "*@temuemail.com OR *@temu.com OR temu@commerce.temuemail.com"
    },
    "Facebook Bildirimleri": {
        "query": "from:facebookmail.com -in:trash",
        "filter_from": "*@facebookmail.com"
    },
    "Trendyol Pazarlama": {
        "query": "from:info@trendyolmail.com -in:trash",
        "filter_from": "info@trendyolmail.com"
    },
    "iyzico Kampanya": {
        "query": "from:duyuru@e.iyzico.com -in:trash",
        "filter_from": "duyuru@e.iyzico.com"
    },
    "Garanti BBVA Bonus Kampanya": {
        "query": "from:bonuscard@kampanya.garantibbva.com.tr -in:trash",
        "filter_from": "bonuscard@kampanya.garantibbva.com.tr"
    },
    "Ziraat Bankkart Reklam": {
        "query": "from:bankkart@yayin.ziraatbank.com.tr -in:trash",
        "filter_from": "bankkart@yayin.ziraatbank.com.tr"
    },
    "MediaMarkt CLUB": {
        "query": "from:info@mail.my.mediamarkt.com.tr -in:trash",
        "filter_from": "info@mail.my.mediamarkt.com.tr"
    },
    "Bolnut": {
        "query": "from:no-reply@bolnut.com -in:trash",
        "filter_from": "no-reply@bolnut.com"
    }
}

def safe_api_call(fn, *args, **kwargs):
    max_retries = 6
    backoff = 10
    for attempt in range(max_retries):
        try:
            return fn(*args, **kwargs)
        except HttpError as e:
            if e.resp.status in [403, 429] and "rateLimitExceeded" in str(e):
                print(f"    [!] Gmail kota sınırı koruması: {backoff} saniye bekleniyor...", flush=True)
                time.sleep(backoff)
                backoff = min(backoff * 2, 60)
            else:
                raise
    return None

def fetch_all_message_ids(gmail, query):
    ids = []
    page_token = None
    while True:
        def _list():
            return gmail.users().messages().list(
                userId='me',
                q=query,
                maxResults=500,
                pageToken=page_token
            ).execute()
        res = safe_api_call(_list)
        if not res:
            break
        msgs = res.get('messages', [])
        ids.extend([m['id'] for m in msgs])
        page_token = res.get('nextPageToken')
        if not page_token:
            break
        time.sleep(0.5)
    return ids

def trash_messages_in_batches(gmail, msg_ids, batch_size=40):
    total = len(msg_ids)
    trashed = 0
    for i in range(0, total, batch_size):
        chunk = msg_ids[i:i + batch_size]
        def _exec_batch():
            batch = gmail.new_batch_http_request()
            for mid in chunk:
                batch.add(gmail.users().messages().trash(userId='me', id=mid))
            batch.execute()
        safe_api_call(_exec_batch)
        trashed += len(chunk)
        print(f"    -> {trashed}/{total} taşındı...", flush=True)
        time.sleep(1.0)
    return trashed

def main():
    store = get_credential_store()
    creds = store.get_credential('ramazaneryildiz@gmail.com')
    gmail = build('gmail', 'v1', credentials=creds)

    print("=" * 65)
    print(" GMAIL KAPSAMLI TEMİZLİK VE FİLTRELEME İŞLEMİ")
    print("=" * 65)

    # Önce Filtreleri kuralım (Kota maliyeti çok düşüktür)
    print("\n--- 1. İLERİYE DÖNÜK FİLTRELER KONTROL EDİLİYOR / OLUŞTURULUYOR ---")
    existing_filters = gmail.users().settings().filters().list(userId='me').execute().get('filter', [])
    existing_froms = [f.get('criteria', {}).get('from', '') for f in existing_filters]

    filters_created = []
    for name, config in TARGET_CATEGORIES.items():
        f_from = config["filter_from"]
        if any(f_from in ef for ef in existing_froms):
            print(f"[i] '{name}' için filtre zaten aktif.")
            continue
        try:
            res = gmail.users().settings().filters().create(
                userId='me',
                body={
                    'criteria': {'from': f_from},
                    'action': {'addLabelIds': ['TRASH']}
                }
            ).execute()
            print(f"[+] '{name}' için filtre oluşturuldu: {f_from}")
            filters_created.append(name)
        except Exception as e:
            print(f"[-] '{name}' filtresi eklenirken hata: {e}")

    # Sonra kalan mailleri temizleyelim
    print("\n--- 2. MEVCUT MAİLLER ÇÖP KUTUSUNA TAŞINIYOR ---")
    stats = {}
    total_trashed_all = 0

    for name, config in TARGET_CATEGORIES.items():
        q = config["query"]
        print(f"\n[+] '{name}' kontrol ediliyor...")
        msg_ids = fetch_all_message_ids(gmail, q)
        count = len(msg_ids)
        print(f"    Çöp kutusuna taşınacak mail sayısı: {count}")
        if count > 0:
            trashed = trash_messages_in_batches(gmail, msg_ids)
            stats[name] = trashed
            total_trashed_all += trashed
        else:
            stats[name] = 0

    print("\n" + "=" * 65)
    print(" İŞLEM BAŞARIYLA TAMAMLANDI")
    print("=" * 65)
    for name, count in stats.items():
        print(f" * {name}: {count} adet mail Çöp Kutusu'na taşındı.")
    print(f"\nBu oturumda temizlenen: {total_trashed_all} adet")
    print(f"(Önceki adımda 989 adet Temu maili taşınmıştı)")
    print(f"Oluşturulan yeni filtre sayısı: {len(filters_created)}")
    print("=" * 65)

if __name__ == "__main__":
    main()

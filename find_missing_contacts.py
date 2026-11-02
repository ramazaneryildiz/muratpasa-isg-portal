import sqlite3

conn = sqlite3.connect(r"c:\isg\isg_muratpasa.db")
c = conn.cursor()

c.execute("SELECT id, kurum_kodu, kurum_adi, bagli_birim, kurum_statusu FROM schools WHERE kurum_statusu = 'Özel' AND bagli_birim != 'Özel Öğretim Kurumları Genel Müdürlüğü'")
rows = c.fetchall()
print(f"Devlet okulu olup Özel görünen kurum sayısı: {len(rows)}")
for r in rows:
    print(r)

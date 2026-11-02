import os, sqlite3, openpyxl, unicodedata

def clean_phone(val):
    if not val: return ''
    s = str(val).split('.')[0].strip()
    digits = ''.join(c for c in s if c.isdigit())
    if len(digits) == 10 and digits.startswith('5'):
        return '0' + digits
    elif len(digits) == 11 and digits.startswith('05'):
        return digits
    elif digits:
        return digits
    return ''

def normalize(s):
    if not s: return ''
    s = str(s).lower().replace('ı', 'i').replace('ğ', 'g').replace('ü', 'u').replace('ş', 's').replace('ö', 'o').replace('ç', 'c')
    s = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in s if c.isalnum() or c.isspace()).strip()

conn = sqlite3.connect(r"c:\isg\isg_muratpasa.db")
c = conn.cursor()

# 1. Fix misclassified state schools
c.execute("UPDATE schools SET kurum_statusu = 'Resmî' WHERE bagli_birim = 'Özel Eğitim ve Rehberlik Hizmetleri Genel Müdürlüğü'")

# 2. Update Ülkü-Seyfi Kandemir Ortaokulu (708801)
c.execute("""
    UPDATE schools
    SET mudur_ad_soyad = 'Yunus VAR',
        mudur_tel = '05356936086',
        isveren_vekili_ad_soyad = 'Fırat ILGIN',
        isveren_vekili_tel = '05546927744'
    WHERE kurum_kodu = '708801'
""")

# 3. Update Ramazan-Hatice Savaş II. Kademe (747613) and III. Kademe (747833)
c.execute("""
    UPDATE schools
    SET mudur_ad_soyad = 'Davut YILDIZ',
        mudur_tel = '05055238377',
        isveren_vekili_ad_soyad = 'Savaş YILDIRAK',
        isveren_vekili_tel = '05053148644'
    WHERE kurum_kodu IN ('747613', '747833')
""")

conn.commit()

# Check all Resmî schools missing phone
c.execute("SELECT id, kurum_kodu, kurum_adi FROM schools WHERE kurum_statusu = 'Resmî' AND (mudur_tel IS NULL OR mudur_tel = '')")
missing_official = c.fetchall()
print(f"Resmî okullardan telefonu eksik kalan: {len(missing_official)}")
for m in missing_official:
    print("  ", m)

# Total with phones
total_with_phone = c.execute("SELECT COUNT(*) FROM schools WHERE mudur_tel IS NOT NULL AND mudur_tel != ''").fetchone()[0]
total_official = c.execute("SELECT COUNT(*) FROM schools WHERE kurum_statusu = 'Resmî'").fetchone()[0]
print(f"\nToplam Resmî Kurum Sayısı: {total_official}")
print(f"Telefonu Olan Okul Sayısı : {total_with_phone}")

conn.close()

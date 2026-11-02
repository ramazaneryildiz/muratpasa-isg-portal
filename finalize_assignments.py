import sqlite3

conn = sqlite3.connect(r"c:\isg\isg_muratpasa.db")
c = conn.cursor()

# Set unassigned schools to NULL
c.execute("""
    UPDATE schools
    SET sorumlu_uzman_id = NULL
    WHERE sorumlu_uzman_id = 1 AND kurum_kodu NOT IN ('969922', '970151', '751351', '970149', '764955', '969926')
""")
conn.commit()

c.execute("""
    SELECT u.id, u.username, u.full_name, u.igu_sinifi, u.kadro_yeri,
           COUNT(s.id) as okul_sayisi,
           COALESCE(SUM(s.personel_sayisi), 0) as toplam_personel,
           COALESCE(SUM(s.aylik_hizmet_suresi_dk), 0) as toplam_dk
    FROM users u
    LEFT JOIN schools s ON u.id = s.sorumlu_uzman_id
    GROUP BY u.id
    ORDER BY okul_sayisi DESC
""")

print("=" * 80)
print(f"{'ID':2} | {'Kullanıcı Adı':16} | {'Adı Soyadı':22} | {'Sınıf':10} | {'Okul':4} | {'Personel':8} | {'Hizmet Süresi':12}")
print("=" * 80)
for r in c.fetchall():
    uid, uname, fname, sinif, kadro, okuls, pers, dk = r
    print(f"{uid:2} | {uname:16} | {fname:22} | {sinif or '':10} | {okuls:4} | {pers:8} | {dk:6} dk ({dk//60} sa)")
print("=" * 80)

cnt_unassigned = c.execute("SELECT COUNT(*) FROM schools WHERE sorumlu_uzman_id IS NULL").fetchone()[0]
cnt_assigned = c.execute("SELECT COUNT(*) FROM schools WHERE sorumlu_uzman_id IS NOT NULL").fetchone()[0]
print(f"Toplam Atanan Resmî Okul Sayısı: {cnt_assigned}")
print(f"Genel Denetim / Özel Kurum Sayısı : {cnt_unassigned}")

conn.close()

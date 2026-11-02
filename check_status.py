import sqlite3

conn = sqlite3.connect(r"c:\isg\isg_muratpasa.db")
c = conn.cursor()

print("=" * 60)
print(" GÜNCEL VERİTABANI İSTATİSTİKLERİ")
print("=" * 60)
print("Toplam Okul Sayısı                :", c.execute("SELECT COUNT(*) FROM schools").fetchone()[0])
print("Müdür / İletişim Telefonu Olan    :", c.execute("SELECT COUNT(*) FROM schools WHERE mudur_tel IS NOT NULL AND mudur_tel != ''").fetchone()[0])
print("Form Yanıtı / Güncel Girişi Olan  :", c.execute("SELECT COUNT(*) FROM schools WHERE son_veri_giris_tarihi IS NOT NULL").fetchone()[0])
print("Emniyet / Kolluk Puanı Olan       :", c.execute("SELECT COUNT(*) FROM schools WHERE toplam_guvenlik_puani IS NOT NULL").fetchone()[0])
print("Sığınak Verisi (Var) Olan         :", c.execute("SELECT COUNT(*) FROM schools WHERE siginak_var_mi = 'Var'").fetchone()[0])
print("Sıfır Atık Belgesi Olan           :", c.execute("SELECT COUNT(*) FROM inspection_metrics WHERE sifir_atik_belge_no IS NOT NULL AND sifir_atik_belge_no != ''").fetchone()[0])
print("Kamera Detayı (İç/Dış) Olan       :", c.execute("SELECT COUNT(*) FROM inspection_metrics WHERE kamera_bina_ici_calisan > 0 OR kamera_bina_disi_calisan > 0").fetchone()[0])
print("Güvenlik İzin Belgesi 'VAR' Olan  :", c.execute("SELECT COUNT(*) FROM inspection_metrics WHERE guvenlik_izin_belgesi_var_mi = 'VAR'").fetchone()[0])
print("El / Kapı Dedektörü 'Var' Olan    :", c.execute("SELECT COUNT(*) FROM inspection_metrics WHERE el_dedektoru_var_mi = 'Var' OR kapi_dedektoru_var_mi = 'Var'").fetchone()[0])
print("İSG Check-List Toplam Madde Sayısı:", c.execute("SELECT COUNT(*) FROM school_isg_checklists").fetchone()[0])
print("=" * 60)

conn.close()

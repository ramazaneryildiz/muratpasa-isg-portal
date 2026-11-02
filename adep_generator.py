import io
import math
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def calculate_quotas(school_type, staff_count):
    staff = max(staff_count or 0, 20)
    is_tehlikeli = (school_type == 'MTAL')
    
    if is_tehlikeli:
        sondurme = max(1, math.ceil(staff / 40))
        kurtarma = max(1, math.ceil(staff / 40))
        koruma = max(1, math.ceil(staff / 40))
        ilkyardim = max(1, math.ceil(staff / 15))
        gecerlilik = "4 Yıl (Tehlikeli Sınıf — 6331 s. Kanun)"
        oran_yangin = "Her 40 çalışana en az 1 kişi (1/40)"
        oran_ilkyardim = "Her 15 çalışana en az 1 kişi (1/15)"
        tehlike_str = "TEHLİKELİ (NACE: 85.32)"
    else:
        sondurme = max(1, math.ceil(staff / 50))
        kurtarma = max(1, math.ceil(staff / 50))
        koruma = max(1, math.ceil(staff / 50))
        ilkyardim = max(1, math.ceil(staff / 20))
        gecerlilik = "6 Yıl (Az Tehlikeli Sınıf — Her eğitim yılı başında EK-7 ile güncellenir)"
        oran_yangin = "Her 50 çalışana en az 1 kişi (1/50)"
        oran_ilkyardim = "Her 20 çalışana en az 1 kişi (1/20)"
        tehlike_str = "AZ TEHLİKELİ (NACE: 85.31)"

    return {
        'is_tehlikeli': is_tehlikeli,
        'tehlike_str': tehlike_str,
        'gecerlilik': gecerlilik,
        'sondurme': sondurme,
        'kurtarma': kurtarma,
        'koruma': koruma,
        'ilkyardim': ilkyardim,
        'oran_yangin': oran_yangin,
        'oran_ilkyardim': oran_ilkyardim,
        'toplam_ekip': sondurme + kurtarma + koruma + ilkyardim
    }

def get_architectural_units(kademe):
    if kademe == 'MTAL':
        return [
            ("1. Torna-Tesviye ve Talaşlı İmalat Atölyesi", "Acil stop butonları, talaş koruyucuları, tezgah arası 120cm emniyet koridoru, topraklama hatları."),
            ("2. CNC ve Bilgisayarlı İmalat Atölyesi", "Elektronik kapı kilit sviçleri, soğutma sıvısı havalandırması, otomatik yangın algılama."),
            ("3. Motorlu Araçlar ve Hidrolik Lift Alanı", "Mekanik kilitli araç kaldırma liftleri, egzoz tahliye aspiratörü, kuyu koruma kapakları."),
            ("4. Otomotiv Boya ve Kurutma Fırın Kabini", "Ex-Proof aydınlatma ve motorlar, statik topraklama baraları, basınç tahliye klapeleri."),
            ("5. Kaynak ve Sıcak İşlem Metal Atölyesi", "Gaz tüpleri zincirli sabitleme, alev geri tepme ventilleri, kaynak dumanı lokal emiş davlumbazları."),
            ("6. Elektrik-Elektronik ve Yüksek Gerilim Lab.", "Kaçak akım röleleri (30mA), yalıtkan çalışma paspasları, şeffaf koruyucu pano kapakları."),
            ("7. Mekatronik, Robotik ve Pnömatik Laboratuvarı", "Pnömatik basınç regülatörleri, robot çalışma hücresi ışık bariyerleri, acil kesiciler."),
            ("8. Bilişim Ağ Altyapısı ve Sunucu Güvenlik Odası", "FM200 / Novec gazlı yangın söndürme, hassas klima iklimlendirme, biyometrik giriş."),
            ("9. Kimya Teknolojileri ve Analitik Cihazlar Lab.", "Asit-baz kimyasal dolapları, acil göz ve vücut duşları, patlamaya dayanıklı çeker ocak."),
            ("10. İnşaat Teknolojileri ve Ahşap Yapı Atölyesi", "Daire testere ayırıcı bıçak ve koruyucu siperlikleri, talaş filtre siklonu, KKD zimmet alanı."),
            ("11. Grafik Tasarım ve Matbaa/Baskı Birimi", "Giyotin bıçak çift el kumanda emniyeti, kimyasal mürekkep MSDS güvenlik kartları."),
            ("12. Yiyecek İçecek Hizmetleri Endüstriyel Mutfak", "Otomatik davlumbaz söndürme (Ansul), gaz dedektörü ve solenoid vana, kaymaz hijyenik zemin.")
        ]
    elif kademe == 'RAM':
        return [
            ("1. Bireysel Değerlendirme ve Görüşme Odaları", "Akustik ses yalıtımı, kırılmaz emniyetli camlar, acil durum çağrı butonu, yumuşak kenarlı mobilya."),
            ("2. Eğitsel Tanılama ve Psikolojik Test Odaları", "Standart aydınlatma ve ısı, dikkat dağıtmayan nötr renkler, sessiz ortam standartları."),
            ("3. Çok Disiplinli Değerlendirme Ekibi (ÇDE) Salonu", "Multidisipliner kurul oturma düzeni, gizlilik esaslı ses izolasyonu, sunum altyapısı."),
            ("4. TS 9111 Uyumlu Engelsiz Giriş Rampaları", "Maksimum %6 eğimli tekerlekli sandalye rampaları, çift taraflı paslanmaz tutunma küpeşteleri."),
            ("5. Engelli Asansörü ve Kat İçi Dikey Taşıma", "Sesli kat anonsu, kabartmalı Braille alfabeli butonlar, acil kurtarma akü sistemi."),
            ("6. Acil İpli Çağrı Sistemli Engelsiz Tuvaletler", "Klozet yanında zemine kadar uzanan acil çağrı ipi, sesli ve ışıklı koridor ikaz lambası, mafsallı tutunma barları."),
            ("7. PİY ve TS-GTM Saha Koordinasyon Odası", "Psikolojik İlk Yardım kriz müdahale çantaları, saha sevk formları, mobil iletişim cihazları."),
            ("8. KVKK Uyumlu Özel Nitelikli Test Arşivi", "Kilitli yangına dayanıklı çelik arşiv dolapları, yetkilendirilmiş personel erişim günlüğü."),
            ("9. İdari Birimler ve Veli Bekleme Salonu", "Geniş koridorlar, engelli şarj istasyonu, bilgilendirme ekranları ve acil çıkış yönlendirmesi."),
            ("10. Kazan Dairesi, Teknik Hacimler ve Sığınak", "Yetkisiz girişe kilitli kapılar, ex-proof aydınlatma, periyodik muayene etiketli basınçlı kaplar.")
        ]
    elif kademe == 'ANADOLU_LISESI':
        return [
            ("1. Genel Derslikler ve Koridorlar", "10-15 cm açılabilen pencere emniyet kilitleri, duvara sabitlenmiş dolaplar, geniş kaçış yolları."),
            ("2. Fizik, Kimya ve Biyoloji Fen Laboratuvarları", "Kilitli kimyasal deposu, çeker ocak, kimyasal döküntü kiti, yangın battaniyesi, güvenlik talimatları."),
            ("3. Bilişim Teknolojileri ve Kodlama Sınıfı", "Korumalı kablo kanalları, kaçak akım koruma sistemi, ergonomik çalışma koltukları."),
            ("4. Görsel Sanatlar ve Müzik Atölyeleri", "Tiner ve boya buharı doğal havalandırması, kesici alet kilitli muhafaza kutusu."),
            ("5. Kapalı Spor Salonu ve Kondisyon Alanı", "Darbe emici koruma minderleri, tribün korkulukları, basketbol potası güvenlik kilitleri."),
            ("6. Z-Kütüphane ve Konferans/Çok Amaçlı Salon", "Panik barlı acil çıkış kapıları, duman dedektörleri, acil aydınlatma ve yönlendirme armatürleri."),
            ("7. Okul Kantini ve Yemekhane Birimi", "72 saatlik numune saklama dolabı, personel hijyen belgeleri, kaymaz zemin paspasları."),
            ("8. Okul Bahçesi, Açık Spor Sahaları ve Nizamiye", "Düzenli Ziyaretçi Kayıt Sistemi, motorlu taşıt giriş engeli bariyeri, aydınlatma direkleri koruması."),
            ("9. Kazan Dairesi, Su Deposu ve Sığınak Alanı", "Doğalgaz alarm dedektörü ve emniyet ventili, hidrofor muayene belgesi, havalandırma menfezleri."),
            ("10. Pansiyon ve Yatakhaneler (Varsa)", "Yangın kaçış merdiveni, duman tahliye kapağı, elektrik panoları kilitli ve yangın söndürücü hazır.")
        ]
    else: # İlkokul / Ortaokul / Anaokulu / Diğer
        return [
            ("1. Temel Eğitim Derslikleri ve Etüt Odaları", "Radyatör koruma kafesleri, priz emniyet kapakları, parmak sıkışma önleyici kapı fitilleri."),
            ("2. Fen ve Teknoloji Dersliği", "Güvenli deney malzemeleri, ilk yardım kiti, kilitli malzeme dolapları."),
            ("3. Özel Eğitim ve Destek Eğitim Odası", "Yumuşak zemin kaplaması, düşme önleyici minderler, pedagojik güvenlik donanımları."),
            ("4. Çok Amaçlı Salon ve Gösteri Alanı", "Sabit sahne basamakları, yangın algılama dedektörleri, acil kaçış levhaları."),
            ("5. Okul Bahçesi ve Çocuk Oyun Alanları", "Kauçuk darbe emici zemin, oyun parkı periyodik kontrolleri, bahçe ihata duvarı korkulukları."),
            ("6. Okul Girişi ve Güvenlik Nizamiye Alanı", "Ziyaretçi kabul bankosu, turnike / kontrollü geçiş kapısı, güvenlik kamerası kör noktasız kayıt."),
            ("7. Kantin ve Beslenme Alanı", "Gıda güvenliği standartları, sıcak su haşlanma önleyici termostatik bataryalar."),
            ("8. Tuvaletler ve Lavabo Hijyen İstasyonları", "Kırılmaz aynalar, fotoselli sıvı sabunluklar, pedallı çöp kovaları, kaymaz zemin."),
            ("9. Kazan Dairesi, Teknik Odalar ve Sığınak", "Sürekli kilitli kapılar, periyodik kontrol etiketli kalorifer kazanı ve yangın tüpleri.")
        ]

def generate_adep_docx(school, metrics=None):
    s = dict(school) if school else {}
    m = dict(metrics) if metrics else {}
    
    k_kodu = str(s.get('kurum_kodu', ''))
    k_adi = s.get('kurum_adi', 'Okul / Kurum Adı')
    kademe = s.get('kademe') or 'ORTAOKUL'
    mudurluk = s.get('genel_mudurluk') or s.get('bagli_birim') or 'Temel Eğitim Genel Müdürlüğü'
    staff_count = s.get('personel_sayisi') or ((s.get('ogretmen_sayisi') or 0) + (s.get('bahce_gorevlisi_sayisi') or 0) + 10)
    
    quotas = calculate_quotas(kademe, staff_count)
    units = get_architectural_units(kademe)

    doc = Document()
    
    # Sayfa Kenar Boşlukları (Standart Resmi Yazı Boyutu: 2.5 cm)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Başlık Alanı
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_top.add_run("T.C.\nANTALYA VALİLİĞİ — MURATPAŞA KAYMAKAMLIĞI\nİlçe Millî Eğitim Müdürlüğü\n")
    r1.bold = True
    r1.font.size = Pt(12)
    r1.font.color.rgb = RGBColor(15, 23, 42)

    r_sch = p_top.add_run(f"{k_adi.upper()}\n(MEBBİS Kurum Kodu: {k_kodu})\n\n")
    r_sch.bold = True
    r_sch.font.size = Pt(13)
    r_sch.font.color.rgb = RGBColor(30, 58, 138)

    r2 = p_top.add_run("ACİL DURUM EYLEM PLANI (ADEP) VE İSG İÇ YÖNERGESİ\n")
    r2.bold = True
    r2.font.size = Pt(14)
    r2.font.color.rgb = RGBColor(185, 28, 28)

    r3 = p_top.add_run("MEB 2024/56 Sayılı Şiddetin Önlenmesi Genelgesi, 6331 Sayılı İSG Kanunu ve MEB Acil Durumlar Yönetmeliği Tam Uyumlu\n(2026–2027 Eğitim ve Öğretim Yılı Resmî Nüshası)\n")
    r3.italic = True
    r3.font.size = Pt(10)
    r3.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 1. BÖLÜM: KURUMSAL KÜNYE TABLOSU
    h1 = doc.add_heading("1. KURUMSAL KİMLİK VE MEVZUAT MATRİSİ", level=1)
    h1.runs[0].font.color.rgb = RGBColor(30, 58, 138)

    table_meta = doc.add_table(rows=8, cols=2)
    table_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_meta.autofit = False

    meta_data = [
        ("Okul / Kurum Adı ve Kodu", f"{k_adi} — Kurum Kodu: {k_kodu}"),
        ("Bağlı Genel Müdürlük", mudurluk),
        ("Kurum Türü ve Kademesi", f"{kademe} ({s.get('kurum_turu', '')})"),
        ("İSG Tehlike Sınıfı & NACE", quotas['tehlike_str']),
        ("Güvenlik Öncelik Derecesi", f"{s.get('oncelik_derecesi', 3)}. Derece Öncelikli Okul (Tablo-1)"),
        ("Toplam Çalışan Sayısı", f"{staff_count} Çalışan (Öğretmen, İdari, TYP/Güvenlik ve Destek Personeli)"),
        ("İşveren ve İşveren Vekili", f"Okul Müdürü: {s.get('mudur_ad_soyad') or 'Müdür'} | İSG Md. Yrd: {s.get('isveren_vekili_ad_soyad') or 'İSG Md. Yrd.'}"),
        ("İSG Koordinasyonu ve Geçerlilik", f"Ramazan ERYILDIZ (A Sınıfı İSG Uzmanı) | Geçerlilik: {quotas['gecerlilik']}")
    ]

    for i, (k, v) in enumerate(meta_data):
        row = table_meta.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
        c0.text = k
        c1.text = v
        c0.paragraphs[0].runs[0].font.bold = True
        c0.paragraphs[0].runs[0].font.size = Pt(9.5)
        c1.paragraphs[0].runs[0].font.size = Pt(9.5)
        set_cell_background(c0, "F1F5F9")
        set_cell_margins(c0)
        set_cell_margins(c1)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 2. BÖLÜM: YASAL KOTA VE ACİL DURUM EKİPLERİ HESAPLAMASI
    h2 = doc.add_heading("2. YASAL ACİL DURUM EKİPLERİ VE PERSONEL KOTA DAĞILIMI", level=1)
    h2.runs[0].font.color.rgb = RGBColor(30, 58, 138)

    p_quota_exp = doc.add_paragraph()
    p_quota_exp.add_run(f"6331 sayılı İSG Kanunu ve MEB Acil Durumlar Yönetmeliği uyarınca; okulumuzun tehlike sınıfı ({quotas['tehlike_str']}) ve toplam çalışan sayısı ({staff_count}) dikkate alınarak kurulması zorunlu asgari acil durum ekipleri aşağıdaki kotalara göre tefrik edilmiştir:")

    table_quota = doc.add_table(rows=6, cols=4)
    table_quota.alignment = WD_TABLE_ALIGNMENT.CENTER

    q_headers = ["Acil Durum Ekibi Adı", "Yasal Görevlendirme Oranı", "Asgari Zorunlu Kişi", "Görev ve Eylem Alanı"]
    hdr_row = table_quota.rows[0]
    for idx, h_text in enumerate(q_headers):
        cell = hdr_row.cells[idx]
        cell.text = h_text
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9.5)
        set_cell_background(cell, "1E3A8A")
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        set_cell_margins(cell)

    q_rows = [
        ("Söndürme Ekibi", quotas['oran_yangin'], f"{quotas['sondurme']} Kişi", "Yangın anında ilk müdahale, itfaiye gelene kadar söndürme/sınırlama"),
        ("Kurtarma Ekibi", quotas['oran_yangin'], f"{quotas['kurtarma']} Kişi", "Öğrenci ve personeli acil toplanma alanına tahliye, arama-kurtarma"),
        ("Koruma Ekibi", quotas['oran_yangin'], f"{quotas['koruma']} Kişi", "Panik ve kargaşayı önleme, okul çevresi emniyeti, evrak ve kayıt emniyeti"),
        ("İlk Yardım Ekibi", quotas['oran_ilkyardim'], f"{quotas['ilkyardim']} Kişi (Sertifikalı)", "Yaralılara standart ilk yardım müdahalesi, 112 koordinasyonu"),
        ("TOPLAM ACİL KADRO", "Tam Teşkilatlı Ekip", f"{quotas['toplam_ekip']} Görevli", "Okul İSG Kurulu koordinasyonunda kriz ve tatbikat yönetimi")
    ]

    for r_idx, r_vals in enumerate(q_rows):
        row = table_quota.rows[r_idx + 1]
        for c_idx, val in enumerate(r_vals):
            cell = row.cells[c_idx]
            cell.text = val
            p = cell.paragraphs[0]
            p.runs[0].font.size = Pt(9)
            if r_idx == 4:
                p.runs[0].font.bold = True
                set_cell_background(cell, "EFF6FF")
            elif r_idx % 2 == 0:
                set_cell_background(cell, "F8FAFC")
            set_cell_margins(cell)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 3. BÖLÜM: OKUL TÜRÜNE ÖZGÜ MİMARİ BİRİMLER VE RİSK TEDBİRLERİ
    h3 = doc.add_heading(f"3. OKULA ÖZGÜ MİMARİ BİRİMLER VE RİSK KONTROL TEDBİRLERİ ({kademe})", level=1)
    h3.runs[0].font.color.rgb = RGBColor(30, 58, 138)

    p_units_desc = doc.add_paragraph()
    p_units_desc.add_run(f"Okulun fiziki yerleşkesi, atölye ve birim yapısına ({kademe}) göre özel olarak tanımlanan risk ve güvenlik kontrol noktaları:")

    for title, desc in units:
        p_u = doc.add_paragraph()
        r_ut = p_u.add_run(f"• {title}: ")
        r_ut.bold = True
        r_ut.font.size = Pt(10)
        r_ud = p_u.add_run(desc)
        r_ud.font.size = Pt(9.5)
        p_u.paragraph_format.space_after = Pt(3)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 4. BÖLÜM: MEB 2024/56 GENELGESİ AKTİF SALDIRGAN VE ŞİDDETİ ÖNLEME PROTOKOLLERİ
    h4 = doc.add_heading("4. MEB 2024/56 ŞİDDETİN ÖNLENMESİ VE KRİZ PROTOKOLLERİ", level=1)
    h4.runs[0].font.color.rgb = RGBColor(185, 28, 28)

    # A. Lockdown Prosedürü
    doc.add_heading("4.1. 'Kilit-Kapıyı Kilitle' (Lockdown) 8 Adımlı Standart Operasyon Prosedürü", level=2)
    lockdown_steps = [
        ("1. Güvenli Alana Giriş:", "Koridordaki ve bahçedeki tüm öğrenci/personel derhal en yakın kapısı kilitlenebilir odaya veya sınıfa alınır."),
        ("2. Kapıyı Kilitleme ve Barikat Kurma:", "Sınıf kapısı kilitlenir; kilit yoksa öğretmen masası ve öğrenci sıraları yığılarak ağır barikat kurulur."),
        ("3. Işık ve Görüşü Kapatma:", "Sınıf ışıkları söndürülür, pencereler ve perdeler/panjurlar tamamen kapatılarak dış görüş kesilir."),
        ("4. Alçak Koruma Pozisyonu:", "Kapı ve pencerelerden uzak, ana taşıyıcı duvar diplerinde çömelerek/yere uzanarak alçak hedef pozisyonu alınır."),
        ("5. Tam Sessizlik ve Telefon Emniyeti:", "Sınıf içinde kesinlikle ses çıkarılmaz; cep telefonları tamamen sessize (titreşimsiz) alınır."),
        ("6. Kapıyı Açmama Kuralı:", "Saldırgan kapıyı çalsa, tehdit etse dahi kapı kesinlikle açılmaz; yalnızca resmî kolluk teyidiyle açılır."),
        ("7. Sessiz Konum Bildirimi:", "Öğretmen SMS veya portal üzerinden okul yönetimine ve 112'ye oda numarasını ve kişi sayısını sessizce bildirir."),
        ("8. Güvenli Çıkış:", "Kolluk kuvvetlerinin 'Tehlike Geçti' resmî duyurusu gelene kadar sınıf/oda kesinlikle terk edilmez.")
    ]
    for step_title, step_desc in lockdown_steps:
        p_ls = doc.add_paragraph()
        r_lst = p_ls.add_run(f"   {step_title} ")
        r_lst.bold = True
        r_lst.font.size = Pt(9.5)
        p_ls.add_run(step_desc).font.size = Pt(9.5)
        p_ls.paragraph_format.space_after = Pt(2)

    # B. Run-Hide-Fight
    doc.add_heading("4.2. 'Kaç - Saklan - Karşı Koy' (Run - Hide - Fight) Hayatta Kalma Modeli", level=2)
    rhf_items = [
        ("1. KAÇ (Run):", "Saldırgandan uzaklaşmak için güvenli bir kaçış koridoru varsa; eşyalarını bırak, hızla güvenli alana koş ve 112'yi ara."),
        ("2. SAKLAN (Hide):", "Kaçış yolu tehlikeliyse; görüş alanından çık, kapıyı kilitle ve barikat kur, ışıkları söndür, yere yat ve sessiz kal."),
        ("3. KARŞI KOY (Fight):", "Yalnızca hayati tehlike altındayken ve son çare olarak; diğer kişilerle organize ol, yangın tüpü/sandalye gibi sert nesnelerle diren.")
    ]
    for r_title, r_desc in rhf_items:
        p_rhf = doc.add_paragraph()
        r_rhft = p_rhf.add_run(f"   {r_title} ")
        r_rhft.bold = True
        r_rhft.font.color.rgb = RGBColor(185, 28, 28)
        p_rhf.add_run(r_desc).font.size = Pt(9.5)

    # C. Sabotaj ve 200m Güvenlik Çemberi
    doc.add_heading("4.3. Sabotaj ve Şüpheli Paket 200 Metre Emniyet Çemberi Protokolü", level=2)
    p_sab = doc.add_paragraph()
    p_sab.add_run("Okul yerleşkesinde sahipsiz çanta, paket veya patlayıcı düzeneği görüldüğünde; kesinlikle dokunulmaz, telsiz/cep telefonu yakınında kullanılmaz. 112 Acil'e bilgi verilerek okul kontrollü tahliye edilir ve kolluk intikal edene kadar binanın ")
    r_sab_bold = p_sab.add_run("en az 200 metre çevresinde emniyet şeridi")
    r_sab_bold.bold = True
    p_sab.add_run(" çekilerek emniyet çemberi oluşturulur.")

    # D. İlk 60 Saniye Matrisi
    doc.add_heading("4.4. Olay Anı — İlk 60 Saniye Kritik Eylemleri Matrisi", level=2)
    table_60 = doc.add_table(rows=5, cols=2)
    table_60.alignment = WD_TABLE_ALIGNMENT.CENTER
    t60_data = [
        ("0 – 15. Saniye", "Tehdidi ilk gören kişi acil alarm butonuna basar, çevredekileri yüksek sesle uyarır."),
        ("15 – 30. Saniye", "Nöbetçi öğretmen ve güvenlik durumu derhal Okul Müdürüne ve idareye bildirir."),
        ("30 – 45. Saniye", "Okul Müdürü / Vekili derhal 112 Acil Çağrı Merkezini arayarak açık adres ve saldırgan bilgisini verir."),
        ("45 – 60. Saniye", "Okul geneli 'Kilit-Kapıyı Kilitle' (Lockdown) anonsu yapılarak tüm birimler korumaya geçer.")
    ]
    t60_hdr = table_60.rows[0]
    t60_hdr.cells[0].text = "Süre / Zaman"
    t60_hdr.cells[1].text = "Uygulanacak Kritik Acil Eylem"
    set_cell_background(t60_hdr.cells[0], "334155")
    set_cell_background(t60_hdr.cells[1], "334155")
    t60_hdr.cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    t60_hdr.cells[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    for i, (tm, act) in enumerate(t60_data):
        row = table_60.rows[i+1]
        row.cells[0].text = tm
        row.cells[1].text = act
        row.cells[0].width = Inches(1.8)
        row.cells[1].width = Inches(4.7)
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        set_cell_margins(row.cells[0])
        set_cell_margins(row.cells[1])

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 5. BÖLÜM: YILLIK GÜVENLİK VE ŞİDDETİ ÖNLEME KONTROL LİSTESİ (12 MADDE)
    h5 = doc.add_heading("5. YILLIK GÜVENLİK VE ŞİDDETİ ÖNLEME KONTROL LİSTESİ (12 MADDE)", level=1)
    h5.runs[0].font.color.rgb = RGBColor(30, 58, 138)

    table_12 = doc.add_table(rows=13, cols=3)
    table_12.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_12.rows[0].cells[0].text = "No"
    table_12.rows[0].cells[1].text = "Kontrol Kriteri ve Standart Şart"
    table_12.rows[0].cells[2].text = "Durum"
    set_cell_background(table_12.rows[0].cells[0], "1E3A8A")
    set_cell_background(table_12.rows[0].cells[1], "1E3A8A")
    set_cell_background(table_12.rows[0].cells[2], "1E3A8A")
    for c in table_12.rows[0].cells:
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        c.paragraphs[0].runs[0].font.bold = True
        set_cell_margins(c)

    c12_items = [
        "Güvenlik kameraları kör nokta bırakmayacak şekilde faal ve kayıtlar en az 30 gün saklanıyor mu?",
        "Acil durum sesli alarm butonları ve uyarı sistemleri test edildi mi ve çalışıyor mu?",
        "Tüm acil çıkış kapıları içeriden panik barlı olup dışarıdan yetkisiz girişe kilitli mi?",
        "Personel acil durum iletişim zinciri ve Okul Kolluk Görevlisi irtibatı güncellendi mi?",
        "Dönemde en az 1 kez 'Kilit-Kapıyı Kilitle' (Lockdown) tatbikatı icra edildi mi?",
        "Yıllık planlı tahliye tatbikatı yapıldı ve toplanma alanında yoklama alındı mı?",
        "Öğrencilere 'Kaç-Saklan-Karşı Koy' ve şiddeti önleme bilgilendirmesi yapıldı mı?",
        "Giriş nizamiyesinde Ziyaretçi Kayıt Defteri ve kimlik kartı uygulaması eksiksiz sürüyor mu?",
        "İlk yardım dolapları ve malzemeleri kontrol edilip eksikleri tamamlandı mı?",
        "Şüpheli/tehlikeli kişi ve paket bildirim talimatı tüm personele yazılı tebliğ edildi mi?",
        "Kriz iletişim planı ve PDR servisi / ÇDE iş akışı hazır mı?",
        "İşbu Acil Durum Planı Eki İSG Kurulu ve Yönetim Ekibince imzalanarak ADEP'e eklendi mi?"
    ]

    for idx, item_txt in enumerate(c12_items):
        row = table_12.rows[idx + 1]
        row.cells[0].width = Inches(0.6)
        row.cells[1].width = Inches(5.0)
        row.cells[2].width = Inches(0.9)
        row.cells[0].text = str(idx + 1)
        row.cells[1].text = item_txt
        row.cells[2].text = "[ X ] UYGUN"
        row.cells[0].paragraphs[0].runs[0].font.size = Pt(8.5)
        row.cells[1].paragraphs[0].runs[0].font.size = Pt(8.5)
        row.cells[2].paragraphs[0].runs[0].font.size = Pt(8.5)
        row.cells[2].paragraphs[0].runs[0].font.bold = True
        if idx % 2 == 1:
            set_cell_background(row.cells[0], "F8FAFC")
            set_cell_background(row.cells[1], "F8FAFC")
            set_cell_background(row.cells[2], "F8FAFC")
        set_cell_margins(row.cells[0])
        set_cell_margins(row.cells[1])
        set_cell_margins(row.cells[2])

    doc.add_paragraph().paragraph_format.space_after = Pt(16)

    # 6. BÖLÜM: YASAL ONAY VE İMZA BLOĞU
    h6 = doc.add_heading("6. YASAL ONAY VE İMZA BLOĞU", level=1)
    h6.runs[0].font.color.rgb = RGBColor(30, 58, 138)

    table_sign = doc.add_table(rows=1, cols=3)
    table_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_sign.autofit = False

    s_c0 = table_sign.rows[0].cells[0]
    s_c1 = table_sign.rows[0].cells[1]
    s_c2 = table_sign.rows[0].cells[2]
    s_c0.width = Inches(2.1)
    s_c1.width = Inches(2.3)
    s_c2.width = Inches(2.1)

    s_c0.text = "PLANI HAZIRLAYAN\n\n\n\n........................................\nİSG'den Sorumlu Md. Yrd.\nİşveren Vekili"
    s_c1.text = "İSG UZMANI KOORDİNASYONU\n\n\n\n........................................\nRamazan ERYILDIZ\nA Sınıfı İş Güvenliği Uzmanı"
    s_c2.text = "ONAYLAYAN\n\n\n\n........................................\nOkul Müdürü\nİşveren (Mühür/İmza)"

    for c in [s_c0, s_c1, s_c2]:
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.size = Pt(8.5)
        p.runs[0].font.bold = True
        set_cell_background(c, "F8FAFC")
        set_cell_margins(c, 180, 180, 120, 120)

    # Memory buffer
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


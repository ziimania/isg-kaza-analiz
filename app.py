import streamlit as st
import pandas as pd
import zipfile
import os
import json
import re
import datetime
import time
import xml.sax.saxutils as saxutils
import google.generativeai as genai

st.set_page_config(page_title="İSG Kök Neden Analizi", layout="wide", page_icon="🛡️")

# Oturum (Session) Hafızası Başlatma
if "islem_tamam" not in st.session_state:
    st.session_state.islem_tamam = False
if "zip_path" not in st.session_state:
    st.session_state.zip_path = ""
if "api_key" not in st.session_state:
    st.session_state.api_key = ""

def extract_json(response_text):
    match = re.search(r'```json\n(.*?)\n```', response_text, re.DOTALL)
    if match: return json.loads(match.group(1))
    try: return json.loads(response_text)
    except: return json.loads(response_text.replace("'", '"'))

def normalize_str(s):
    return str(s).replace('İ','I').replace('i','I').replace('ğ','G').replace('Ğ','G') \
                 .replace('ş','S').replace('Ş','S').replace('ö','O').replace('Ö','O') \
                 .replace('ü','U').replace('Ü','U').replace('ç','C').replace('Ç','C').upper()

def get_val(row, keywords):
    for col in row.index:
        col_norm = normalize_str(col)
        if all(normalize_str(k) in col_norm for k in keywords):
            val = row[col]
            return val if pd.notna(val) else ""
    return ""

def sanitize_text(text):
    if pd.isna(text): return ""
    text = str(text)
    text = re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]', '', text)
    return saxutils.escape(text)

CHECKBOX_MAP = {
  "Ölüm": "xl/ctrlProps/ctrlProp1.xml", "Uzuv Kayıplı": "xl/ctrlProps/ctrlProp2.xml",
  "Tıbbi Müdahale": "xl/ctrlProps/ctrlProp3.xml", "Hafif Yaralanma": "xl/ctrlProps/ctrlProp4.xml",
  "Yaralanmasız Kaza": "xl/ctrlProps/ctrlProp5.xml",
  "Hareketli Aksamlar": "xl/ctrlProps/ctrlProp58.xml", "El Aletleri": "xl/ctrlProps/ctrlProp59.xml",
  "Düşen Malzeme": "xl/ctrlProps/ctrlProp63.xml", "Elle Taşıma": "xl/ctrlProps/ctrlProp65.xml",
  "Yüksekten Düşme": "xl/ctrlProps/ctrlProp67.xml", "Yakalanma / Kaptırma": "xl/ctrlProps/ctrlProp68.xml",
  "Takılma": "xl/ctrlProps/ctrlProp69.xml", "Kayma": "xl/ctrlProps/ctrlProp70.xml",
  "Düşme": "xl/ctrlProps/ctrlProp71.xml", "İki Nesne Arasına Sıkışma": "xl/ctrlProps/ctrlProp74.xml",
  "Lider ve Yönetim": "xl/ctrlProps/ctrlProp36.xml", "Kişisel Koruyucu Donanım": "xl/ctrlProps/ctrlProp37.xml",
  "Görev Analizleri ve Prosedürleri": "xl/ctrlProps/ctrlProp42.xml", "Eksik / yetersiz talimat": "xl/ctrlProps/ctrlProp30.xml",
  "Prosedür ve kuralların takip edilmemesi": "xl/ctrlProps/ctrlProp27.xml",
  "Diğer...": "xl/ctrlProps/ctrlProp88.xml", "Diğer": "xl/ctrlProps/ctrlProp88.xml"
}

HUCRE_X_HARITASI = {
  "Ekipmanı izinsiz kullanmak": "B30", "KKD kullanımında ihmal": "B38", "Prosedür ve kuralları takip etmemek": "B51", "Dikkatsiz çalışma": "B54",
  "Yetersiz veya uygunsuz KKD": "K31", "Tehlikeli çevre koşulları: gazlar, tozlar": "K45", "Yetersiz talimat olması": "K53", "Yetersiz prosedür olması": "K54",
  "Fiziksel uygunsuzluk": "T30", "Zihinsel uygunsuzluk": "T31", "Bilgi eksikliği": "T32", "İşin gerektirdiği tecrübede eksiklik": "T33",
  "Alışkanlıklar": "T34", "Lider eksikliği": "T35", "Görev tanımı yapılmaması": "T36", "İş güvenliği kuralları yetersizliği": "T37",
  "Beceri eksikliği": "AC30", "Fiziksel stres": "AC31", "Zihinsel stres": "AC32", "Yetersiz motivasyon": "AC33",
  "Uygunsuz davranış": "AC34", "Yaptırım yetersizliği": "AC35", "Denetim yetersizliği": "AC36", "Uygun olmayan görevlendirme": "AC37"
}

def safe_update_sheet1(sheet_xml, coord, new_value):
    pattern = rf'<c r="{coord}"(?:[^>]*)?>'
    match = re.search(pattern, sheet_xml)
    if not match: return sheet_xml
    start_tag = match.group(0)
    
    if start_tag.endswith('/>'):
        attrs = start_tag[1:-2]
        end_idx = match.end()
    else:
        attrs = start_tag[1:-1]
        end_match = re.search(r'</c>', sheet_xml[match.end():])
        end_idx = match.end() + end_match.end()
        
    attrs = re.sub(r'\s+t="[^"]*"', '', attrs)
    safe_value = sanitize_text(new_value)
    new_cell = f'<{attrs} t="inlineStr"><is><t>{safe_value}</t></is></c>'
    
    return sheet_xml[:match.start()] + new_cell + sheet_xml[end_idx:]

def update_excel_template(template_path, output_path, updates_dict, labels_to_check):
    files_to_check = [CHECKBOX_MAP[lbl] for lbl in labels_to_check if lbl in CHECKBOX_MAP]
    
    with zipfile.ZipFile(template_path, 'r') as zin:
        sheet_xml = zin.read('xl/worksheets/sheet1.xml').decode('utf-8')
    
    for coord, text_val in updates_dict.items():
        if pd.isna(text_val) or text_val == "": continue
        sheet_xml = safe_update_sheet1(sheet_xml, coord, text_val)
                
    with zipfile.ZipFile(template_path, 'r') as zin, zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'xl/worksheets/sheet1.xml':
                zout.writestr(item.filename, sheet_xml.encode('utf-8'))
            elif item.filename in files_to_check:
                xml_content = zin.read(item.filename).decode('utf-8')
                if 'checked="Checked"' not in xml_content:
                    xml_content = xml_content.replace('<formControlPr ', '<formControlPr checked="Checked" ')
                zout.writestr(item.filename, xml_content.encode('utf-8'))
            else:
                zout.writestr(item.filename, zin.read(item.filename))

MERKEZI_SABLON = "template.xlsx"
os.makedirs("temp_reports", exist_ok=True)

st.title("🛡️ İSG Kök Neden Analizi Otomasyonu")

with st.sidebar:
    st.header("⚙️ Profil ve Güvenli Ayarlar")
    profil_adi = st.text_input("Profil Adınız:", value="")
    
    if profil_adi:
        # Doğrudan oturum hafızasına bağlanan input alanı
        api_input = st.text_input("Gemini API Anahtarınız:", value=st.session_state.api_key, type="password")
        
        if api_input != st.session_state.api_key:
            st.session_state.api_key = api_input.strip()
            
        if st.session_state.api_key:
            st.success("✅ API Anahtarı Aktif ve Kayıtlı")
        else:
            st.warning("⚠️ Lütfen API anahtarınızı girin")
        
        st.markdown("---")
        if os.path.exists(MERKEZI_SABLON):
            st.success("✅ Merkezi Şablon Sistemde Yüklü")
        else:
            st.error("⚠️ Merkezi Şablon (template.xlsx) Bulunamadı!")

if not profil_adi:
    st.info("👈 Lütfen sol menüden Profil Adınızı girerek başlayın.")
else:
    if not os.path.exists(MERKEZI_SABLON):
        st.error("⚠️ Sistemde ana şablon ('template.xlsx') bulunamadı. Lütfen GitHub deposuna bu dosyayı yükleyin.")
    elif not st.session_state.api_key:
        st.warning("👈 Lütfen sol menüden API anahtarınızı girin.")
    else:
        st.write("### 📝 Kaza Verisi Yükle")
        data_file = st.file_uploader("Doldurulmuş Kaza Listesini (Excel / .xlsm) Yükleyin", type=["xlsx", "xlsm"])
        
        if data_file:
            df = pd.read_excel(data_file)
            df.columns = df.columns.str.strip().str.upper()
            df = df.dropna(subset=['ADI SOYADI'])
            
            st.markdown("---")
            st.subheader("🔍 Hızlı Filtreleme Paneli")
            
            col_uzman, col_ay, col_daire = st.columns(3)
            
            uzman_col_name = next((c for c in df.columns if 'UZMAN' in c or 'İSG' in c), None)
            ay_col_name = next((c for c in df.columns if 'AY' in c and ('KAZA' in c or 'TAR' in c)), None)
            daire_col_name = next((c for c in df.columns if 'DAİRE' in c or 'BAĞLI' in c or 'BİRİM' in c), None)
            
            filtered_df = df.copy()
            
            with col_uzman:
                if uzman_col_name:
                    unique_uzmanlar = sorted(df[uzman_col_name].dropna().astype(str).unique())
                    secilen_uzmanlar = st.multiselect("İş Güvenliği Uzmanı", unique_uzmanlar, default=[])
                    if secilen_uzmanlar:
                        filtered_df = filtered_df[filtered_df[uzman_col_name].astype(str).isin(secilen_uzmanlar)]
                else:
                    st.info("İş Güvenliği Uzmanı sütunu bulunamadı.")
                    
            with col_ay:
                if ay_col_name:
                    unique_aylar = sorted(df[ay_col_name].dropna().astype(str).unique())
                    secilen_aylar = st.multiselect("Kaza Tarihi (Ay)", unique_aylar, default=[])
                    if secilen_aylar:
                        filtered_df = filtered_df[filtered_df[ay_col_name].astype(str).isin(secilen_aylar)]
                else:
                    st.info("Kaza Ayı sütunu bulunamadı.")
                    
            with col_daire:
                if daire_col_name:
                    unique_daireler = sorted(df[daire_col_name].dropna().astype(str).unique())
                    secilen_daireler = st.multiselect("Bağlı Bulunduğu Daire", unique_daireler, default=[])
                    if secilen_daireler:
                        filtered_df = filtered_df[filtered_df[daire_col_name].astype(str).isin(secilen_daireler)]
                else:
                    st.info("Daire Başkanlığı sütunu bulunamadı.")
            
            st.markdown("---")
            st.write(f"### 📋 Analiz Edilecek Kazalar (Filtrelenen: {len(filtered_df)} / Toplam: {len(df)})")
            
            selected_indices = []
            
            for idx, row in filtered_df.iterrows():
                isim = row.get('ADI SOYADI', f'Personel_{idx}')
                kaza_turu = str(row.get('KAZA TÜRÜ', ''))
                if st.checkbox(f"{isim} - {kaza_turu}", key=f"chk_{idx}"):
                    selected_indices.append(idx)
                    
            if st.button("🚀 Seçili Raporları Üret", type="primary"):
                if not selected_indices:
                    st.error("Lütfen en az bir kişi seçin!")
                else:
                    st.session_state.islem_tamam = False 
                    
                    with st.spinner('Yapay zeka analizleri gerçekleştiriyor, lütfen bu sayfadan ayrılmayın...'):
                        genai.configure(api_key=st.session_state.api_key)
                        model = genai.GenerativeModel('gemini-3.6-flash')
                        
                        zip_filename = f"temp_reports/ISG_Raporlari_{profil_adi.replace(' ', '_')}.zip"
                        progress_bar = st.progress(0, text="Analiz başlatılıyor...")
                        
                        with zipfile.ZipFile(zip_filename, 'w') as zipf:
                            for i, idx in enumerate(selected_indices):
                                row = df.loc[idx]
                                isim = row.get('ADI SOYADI', f'Personel_{idx}')
                                kaza_turu = str(row.get('KAZA TÜRÜ', ''))
                                
                                progress_bar.progress((i) / len(selected_indices), text=f"Analiz ediliyor: {isim} ({i+1}/{len(selected_indices)})")
                                
                                birim = get_val(row, ['BAĞLI', 'DAİRE'])
                                gorevi = get_val(row, ['GÖREVİ'])
                                ise_giris = get_val(row, ['İŞE', 'GİRİŞ'])
                                if not ise_giris: ise_giris = "-"
                                
                                dogum_gun = get_val(row, ['DOĞUM', 'GÜN'])
                                dogum_ay = get_val(row, ['DOĞUM', 'AY'])
                                dogum_yil = get_val(row, ['DOĞUM', 'YIL'])
                                dogum_tarihi = f"{dogum_gun}.{dogum_ay}.{dogum_yil}" if dogum_gun else ""
                                
                                kaza_gun = get_val(row, ['KAZA', 'TAR', 'GÜN'])
                                kaza_ay = get_val(row, ['KAZA', 'TAR', 'AY'])
                                kaza_yil = get_val(row, ['KAZA', 'TAR', 'YIL'])
                                kaza_tarihi = f"{kaza_gun}.{kaza_ay}.{kaza_yil}" if kaza_gun else ""
                                
                                rapor_tarihi_str = ""
                                try:
                                    k_dt = datetime.datetime(int(kaza_yil), int(kaza_ay), int(kaza_gun))
                                    kayip_is_gunu = get_val(row, ['KAYIP', 'İŞ', 'GÜN'])
                                    try: kayip_gun_int = int(float(kayip_is_gunu))
                                    except: kayip_gun_int = 0
                                    
                                    r_dt = k_dt + datetime.timedelta(days=kayip_gun_int)
                                    r_dt += datetime.timedelta(days=1)
                                    while r_dt.weekday() >= 5: 
                                        r_dt += datetime.timedelta(days=1)
                                    rapor_tarihi_str = r_dt.strftime("%d.%m.%Y")
                                except:
                                    rapor_tarihi_str = datetime.datetime.now().strftime("%d.%m.%Y")
                                
                                updates = {
                                    'A7': isim, 'J7': gorevi, 'R7': dogum_tarihi, 'W7': ise_giris,
                                    'AA7': birim, 'AH2': kaza_tarihi, 'AH3': rapor_tarihi_str
                                }
                                
                                prompt = f"""
                                Sen ÜST DÜZEY bir İş Sağlığı ve Güvenliği (İSG) Uzmanısın ve Kök Neden Analizi konusunda profesyonelsin.
                                Amacın aşağıdaki kaza verilerini irdeleyerek profesyonel bir rapor oluşturmak.
                                
                                Kaza Verileri: {row.to_dict()}
                                Kaza Türü: {kaza_turu}
                                
                                ÖNEMLİ KURALLAR:
                                1. Düzeltici faaliyetleri (DÖF) ASLA destan gibi uzun paragraflar halinde yazma. 
                                2. Her bir DÖF tek bir satıra rahatça sığacak kadar KISA, SADE ve NET bir cümle olmalıdır. 
                                3. C69, C70 ve C71 hücrelerine birbirinden tamamen BAĞIMSIZ ve FARKLI birer faaliyet maddesi yaz.
                                4. KESİNLİKLE maddelerin başına "1.", "2.", "3." gibi rakamlar, tire (-) veya nokta koyma! Şablonda rakamlar zaten var, sadece faaliyet cümlesini yaz.
                                5. SORUMLU KİŞİ (R69, R70, R71): Bu işyeri bir belediye iştirakidir (alt işveren). İştirak firması sadece personel sağlamaktadır. Alet, ekipman temini, tadilat, tamirat, montaj ve personelin sahada görevlendirilmesi tamamen BELEDİYE'nin (Üst İşveren) sorumluluğundadır. Yazdığın DÖF maddesi ekipman temini, bakım, tamirat veya iş/görevlendirme prosedürleri ile ilgiliyse sorumlu alana kesinlikle "Üst İşveren" yaz. Eğer eğitim, evrak takibi, risk analizi gibi İSG süreçleriyse "İSG Birimi" yaz.
                                
                                SADECE JSON döndür. JSON Şablonu:
                                {{
                                    "kaza_sonucu_kutusu": "Ölüm, Uzuv Kayıplı, Tıbbi Müdahale, Hafif Yaralanma, Yaralanmasız Kaza (Kaza türüne ve detaylara bakarak bu 5 seçenekten SADECE BİRİNİ buraya yaz)",
                                    "B9": "Kaza Olay Özeti: Olayın detaylarını anlatan resmi bir İSG açıklaması.",
                                    "L17": "Yaralanmanın vücuttaki tam yeri ve şiddeti.",
                                    "L19": "{kaza_tarihi}",
                                    "L21": "Olayın gerçekleştiği tam nokta/çalışma alanı",
                                    "L23": "Nasıl Gerçekleşti: Olayın oluş şeklini mekanik ve insan faktörlerini dikkate alarak açıkla.",
                                    "L25": "Hangi İş: Olay anında yapılan spesifik işin tam tipi.",
                                    "L27": "{gorevi}",
                                    "isaretlenecek_kutular": ["Seçilen Kutu 1"],
                                    "x_yazilacak_sebepler": ["Sebep 1"],
                                    "C69": "Kısa ve net birinci düzeltici faaliyet. (Başına rakam koyma)",
                                    "R69": "Birinci faaliyetin sorumlusu (Maddenin içeriğine göre İSG Birimi veya Üst İşveren)",
                                    "C70": "Kısa ve net ikinci düzeltici faaliyet. (Gerekiyorsa yaz, yoksa boş bırak. Başına rakam koyma)",
                                    "R70": "İkinci faaliyetin sorumlusu (C70 boş değilse İSG Birimi veya Üst İşveren yaz, boşsa boş bırak)",
                                    "C71": "Kısa ve net üçüncü düzeltici faaliyet. (Gerekiyorsa yaz, yoksa boş bırak. Başına rakam koyma)",
                                    "R71": "Üçüncü faaliyetin sorumlusu (C71 boş değilse İSG Birimi veya Üst İşveren yaz, boşsa boş bırak)"
                                }}
                                
                                'isaretlenecek_kutular' listesine ŞU KELİMELERDEN olaya en uygun olanları (en fazla 4 tane) seç:
                                Hareketli Aksamlar, El Aletleri, Düşen Malzeme, Elle Taşıma, Yüksekten Düşme, Yakalanma / Kaptırma, Takılma, Kayma, Düşme, İki Nesne Arasına Sıkışma, Lider ve Yönetim, Kişisel Koruyucu Donanım, Görev Analizleri ve Prosedürleri, Eksik / yetersiz talimat, Prosedür ve kuralların takip edilmemesi.
                                NOT: Eğer kazaya neden olan etmenler bu listedekilerden hiçbirine uymuyorsa, metin olarak açıklama YAPMA. SADECE "Diğer..." kelimesini listeye ekle.
                                
                                'x_yazilacak_sebepler' listesine ŞU KELİMELERDEN olaya en uygun Görünür ve Temel Sebepleri (en fazla 5 tane) seçip ekle:
                                Ekipmanı izinsiz kullanmak, KKD kullanımında ihmal, Prosedür ve kuralları takip etmemek, Dikkatsiz çalışma, Yetersiz veya uygunsuz KKD, Tehlikeli çevre koşulları: gazlar, tozlar, Yetersiz talimat olması, Yetersiz prosedür olması, Fiziksel uygunsuzluk, Zihinsel uygunsuzluk, Bilgi eksikliği, İşin gerektirdiği tecrübede eksiklik, Alışkanlıklar, Lider eksikliği, Görev tanımı yapılmaması, İş güvenliği kuralları yetersizliği, Beceri eksikliği, Fiziksel stres, Zihinsel stres, Yetersiz motivasyon, Uygunsuz davranış, Yaptırım yetersizliği, Denetim yetersizliği, Uygun olmayan görevlendirme
                                """
                                
                                max_deneme = 3
                                for deneme in range(max_deneme):
                                    try:
                                        response = model.generate_content(prompt)
                                        ai_data = extract_json(response.text)
                                        
                                        kutular = ai_data.pop("isaretlenecek_kutular", [])
                                        
                                        kaza_sonucu = ai_data.pop("kaza_sonucu_kutusu", "")
                                        for key in ["Ölüm", "Uzuv Kayıplı", "Tıbbi Müdahale", "Hafif Yaralanma", "Yaralanmasız Kaza"]:
                                            if key.lower() in str(kaza_sonucu).lower():
                                                kutular.append(key)
                                                break
                                        
                                        x_sebepler = ai_data.pop("x_yazilacak_sebepler", [])
                                        for sebep in x_sebepler:
                                            if sebep in HUCRE_X_HARITASI:
                                                updates[HUCRE_X_HARITASI[sebep]] = "X"
                                                
                                        updates.update(ai_data)
                                        out_name = f"temp_reports/{isim.replace(' ', '_')}_Raporu.xlsx"
                                        update_excel_template(MERKEZI_SABLON, out_name, updates, kutular)
                                        
                                        zipf.write(out_name, arcname=os.path.basename(out_name))
                                        os.remove(out_name)
                                        
                                        time.sleep(6) 
                                        break 
                                        
                                    except Exception as e:
                                        hata_msaji = str(e).lower()
                                        if "429" in hata_msaji or "quota" in hata_msaji or "exhausted" in hata_msaji:
                                            progress_bar.progress((i) / len(selected_indices), text=f"API Limiti doldu, 15sn bekleniyor... ({isim})")
                                            time.sleep(15)
                                        else:
                                            st.error(f"{isim} hatası: {e}")
                                            break 
                            
                            progress_bar.progress(1.0, text="Analiz tamamlandı!")
                            
                        st.session_state.islem_tamam = True
                        st.session_state.zip_path = zip_filename

            if st.session_state.islem_tamam and os.path.exists(st.session_state.zip_path):
                st.success("🎉 Raporlar başarıyla oluşturuldu!")
                with open(st.session_state.zip_path, "rb") as f:
                    st.download_button("📥 Raporları İndir (ZIP)", f, file_name="ISG_Raporlari.zip", mime="application/zip")

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

# --- KULLANICI HAFIZA YÖNETİMİ ---
# Streamlit Cloud'da dosyaları tutmak için 'users_data' klasörü oluşturuyoruz
if not os.path.exists("users_data"):
    os.makedirs("users_data")

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
  "Prosedür ve kuralların takip edilmemesi": "xl/ctrlProps/ctrlProp27.xml"
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
    
    # FORMAT KALKANI: Belirli hücrelerde stili sil (sola yaslı ve ince yazsın diye)
    if coord.startswith('L') or coord.startswith('B9') or coord.startswith('C69') or coord.startswith('C70') or coord.startswith('C71'):
        attrs = re.sub(r'\s+s="[^"]*"', '', attrs)
        
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


# --- ARAYÜZ TASARIMI ---
st.title("🛡️ İSG Kök Neden Analizi Otomasyonu")

# Kullanıcı Kimliği (Cookie mantığı yerine basit profil adı)
with st.sidebar:
    st.header("⚙️ Profil ve Ayarlar")
    profil_adi = st.text_input("Profil Adınız (Sizi hatırlamamız için):", value="")
    
    if profil_adi:
        user_folder = f"users_data/{profil_adi.replace(' ', '_')}"
        if not os.path.exists(user_folder):
            os.makedirs(user_folder)
            
        api_file = f"{user_folder}/api_key.txt"
        template_file = f"{user_folder}/template.xlsx"
        
        # Mevcut API key okuma
        current_api = ""
        if os.path.exists(api_file):
            with open(api_file, "r") as f:
                current_api = f.read()
                
        new_api = st.text_input("Gemini API Anahtarınız:", value=current_api, type="password")
        if new_api != current_api:
            with open(api_file, "w") as f:
                f.write(new_api)
            st.success("API Anahtarı Kaydedildi!")
            
        # Mevcut Şablon durumu
        if os.path.exists(template_file):
            st.success("✅ Şablon Yüklü")
        else:
            st.warning("⚠️ Şablon Yüklü Değil")
            
        uploaded_template = st.file_uploader("Orijinal Boş Şablonu Yükle / Güncelle", type=["xlsx"])
        if uploaded_template:
            with open(template_file, "wb") as f:
                f.write(uploaded_template.getbuffer())
            st.success("Şablon başarıyla kaydedildi!")
            st.rerun()

# ANA EKRAN
if not profil_adi:
    st.info("👈 Lütfen sol menüden Profil Adınızı girerek başlayın.")
else:
    user_folder = f"users_data/{profil_adi.replace(' ', '_')}"
    api_file = f"{user_folder}/api_key.txt"
    template_file = f"{user_folder}/template.xlsx"
    
    if not os.path.exists(api_file) or not os.path.exists(template_file):
        st.warning("👈 Lütfen sol menüden API anahtarınızı ve Boş Şablonunuzu yükleyin.")
    else:
        st.write("### 📝 Kaza Verisi Yükle")
        data_file = st.file_uploader("Doldurulmuş Kaza Listesini (Excel) Yükleyin", type=["xlsx"])
        
        if data_file:
            df = pd.read_excel(data_file)
            df.columns = df.columns.str.strip().str.upper()
            df = df.dropna(subset=['ADI SOYADI'])
            
            st.write("### 🔍 Analiz Edilecek Kazaları Seçin")
            selected_indices = []
            
            # Checkbox tablosu
            for idx, row in df.iterrows():
                isim = row.get('ADI SOYADI', f'Personel_{idx}')
                kaza_turu = row.get('KAZA TÜRÜ', '')
                if st.checkbox(f"{isim} - {kaza_turu}", key=f"chk_{idx}"):
                    selected_indices.append(idx)
                    
            if st.button("🚀 Seçili Raporları Üret", type="primary"):
                if not selected_indices:
                    st.error("Lütfen en az bir kişi seçin!")
                else:
                    with open(api_file, "r") as f:
                        api_key = f.read().strip()
                    
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel('gemini-2.5-flash')
                    
                    zip_filename = "Otomatik_Kok_Neden_Analizleri.zip"
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    with zipfile.ZipFile(zip_filename, 'w') as zipf:
                        for i, idx in enumerate(selected_indices):
                            row = df.loc[idx]
                            isim = row.get('ADI SOYADI', f'Personel_{idx}')
                            status_text.text(f"Analiz ediliyor: {isim} ({i+1}/{len(selected_indices)})")
                            
                            # Veri Okuma
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
                            
                            kaza_turu = str(get_val(row, ['KAZA', 'TÜRÜ'])).lower()
                            kaza_turu_etiket = ""
                            if 'ölüm' in kaza_turu: kaza_turu_etiket = "Ölüm"
                            elif 'uzuv' in kaza_turu: kaza_turu_etiket = "Uzuv Kayıplı"
                            elif 'tıbbi' in kaza_turu or 'ayakta' in kaza_turu or 'tedavi' in kaza_turu: kaza_turu_etiket = "Tıbbi Müdahale"
                            elif 'hafif' in kaza_turu: kaza_turu_etiket = "Hafif Yaralanma"
                            elif 'yaralanmasız' in kaza_turu or 'maddi' in kaza_turu: kaza_turu_etiket = "Yaralanmasız Kaza"
                            
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
                            Amacın aşağıdaki kaza verilerini YÜZEYSEL DEĞİL, çok derinlemesine irdeleyerek profesyonel bir rapor oluşturmak.
                            DİKKAT: Düzeltici faaliyetleri çok uzun bir paragraf yerine, kısa, öz, sade ve anlaşılır cümleler halinde (madde madde) yazmalısın. Eğer faaliyet tek bir satıra sığmayacak kadar uzunsa, C70 ve C71 hücrelerini de kullan.
                            
                            Kaza Verileri: {row.to_dict()}
                            
                            SADECE JSON döndür. JSON Şablonu:
                            {{
                                "B9": "Kaza Olay Özeti: Olayın detaylarını anlatan resmi bir İSG açıklaması.",
                                "L17": "Yaralanmanın vücuttaki tam yeri ve şiddeti.",
                                "L19": "{kaza_tarihi}",
                                "L21": "Olayın gerçekleştiği tam nokta/çalışma alanı",
                                "L23": "Nasıl Gerçekleşti: Olayın oluş şeklini mekanik ve insan faktörlerini dikkate alarak açıkla.",
                                "L25": "Hangi İş: Olay anında yapılan spesifik işin tam tipi.",
                                "L27": "{gorevi}",
                                "isaretlenecek_kutular": ["Seçilen Kutu 1"],
                                "x_yazilacak_sebepler": ["Sebep 1"],
                                "C69": "Düzeltici Önleyici Faaliyet (DÖF) Madde 1: Sade ve anlaşılır kısa bir önlem.",
                                "C70": "Düzeltici Önleyici Faaliyet (DÖF) Madde 2 (Sadece gerekiyorsa, yoksa boş bırak)",
                                "C71": "Düzeltici Önleyici Faaliyet (DÖF) Madde 3 (Sadece gerekiyorsa, yoksa boş bırak)",
                                "R69": "İSG Birimi"
                            }}
                            
                            'isaretlenecek_kutular' listesine ŞU KELİMELERDEN olaya en uygun olanları (en fazla 4 tane) seçip ekle:
                            Hareketli Aksamlar, El Aletleri, Düşen Malzeme, Elle Taşıma, Yüksekten Düşme, Yakalanma / Kaptırma, Takılma, Kayma, Düşme, İki Nesne Arasına Sıkışma, Lider ve Yönetim, Kişisel Koruyucu Donanım, Görev Analizleri ve Prosedürleri, Eksik / yetersiz talimat, Prosedür ve kuralların takip edilmemesi.
                            NOT: Eğer hiçbir kutu kazanın türüne uymuyorsa, listede olmasa bile "Diğer" kelimesini ekle.
                            
                            'x_yazilacak_sebepler' listesine ŞU KELİMELERDEN olaya en uygun Görünür ve Temel Sebepleri (en fazla 5 tane) seçip ekle:
                            Ekipmanı izinsiz kullanmak, KKD kullanımında ihmal, Prosedür ve kuralları takip etmemek, Dikkatsiz çalışma, Yetersiz veya uygunsuz KKD, Tehlikeli çevre koşulları: gazlar, tozlar, Yetersiz talimat olması, Yetersiz prosedür olması, Fiziksel uygunsuzluk, Zihinsel uygunsuzluk, Bilgi eksikliği, İşin gerektirdiği tecrübede eksiklik, Alışkanlıklar, Lider eksikliği, Görev tanımı yapılmaması, İş güvenliği kuralları yetersizliği, Beceri eksikliği, Fiziksel stres, Zihinsel stres, Yetersiz motivasyon, Uygunsuz davranış, Yaptırım yetersizliği, Denetim yetersizliği, Uygun olmayan görevlendirme
                            """
                            
                            max_deneme = 3
                            for deneme in range(max_deneme):
                                try:
                                    response = model.generate_content(prompt)
                                    ai_data = extract_json(response.text)
                                    
                                    kutular = ai_data.pop("isaretlenecek_kutular", [])
                                    if kaza_turu_etiket: kutular.append(kaza_turu_etiket)
                                    
                                    x_sebepler = ai_data.pop("x_yazilacak_sebepler", [])
                                    for sebep in x_sebepler:
                                        if sebep in HUCRE_X_HARITASI:
                                            updates[HUCRE_X_HARITASI[sebep]] = "X"
                                    
                                    updates.update(ai_data)
                                    out_name = f"{isim.replace(' ', '_')}_Raporu.xlsx"
                                    update_excel_template(template_file, out_name, updates, kutular)
                                    
                                    zipf.write(out_name)
                                    os.remove(out_name)
                                    
                                    time.sleep(6) 
                                    break 
                                    
                                except Exception as e:
                                    hata_msaji = str(e).lower()
                                    if "429" in hata_msaji or "quota" in hata_msaji or "exhausted" in hata_msaji:
                                        status_text.text(f"API Limiti doldu, 15 saniye bekleniyor... (Deneme {deneme+1}/{max_deneme})")
                                        time.sleep(15)
                                    else:
                                        st.error(f"{isim} hatası: {e}")
                                        break 
                                        
                            progress_bar.progress((i + 1) / len(selected_indices))
                            
                    status_text.text("🎉 Tüm Raporlar Hazır!")
                    with open(zip_filename, "rb") as f:
                        st.download_button("📥 Raporları İndir (ZIP)", f, file_name="ISG_Raporlari.zip", mime="application/zip")

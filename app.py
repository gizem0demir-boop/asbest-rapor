import streamlit as st
from modules.ayp import render_ayp_module  # Raporlama modülün
from modules.kalite import render_kalite_yonetim_module  # Yeni eklediğimiz kalite modülü

st.set_page_config(
    page_title="Asbest ve Atık Yönetim Rapor Sistemi", layout="wide"
)

# Sol menü - İşlem Kategorisi Seçimi
islem_kategorisi = st.sidebar.selectbox(
    "📂 İşlem Kategorisi Seçin:",
    [
        "-- Seçiniz --",
        "📊 Raporlama İşlemleri",
        "🏗️ Yıkım Planı ve Yasal Evrak Modülü",
        "🧪 ISO/IEC 17025 Kalite Yönetimi",
    ],
)

if islem_kategorisi == "-- Seçiniz --":
    st.markdown("### 🏢 Asbest ve Atık Yönetim Rapor Sistemi")
    st.info(
        "💡 Lütfen sol menüden yapacağınız işlem kategorisini seçerek devam"
        " edin."
    )

elif islem_kategorisi == "📊 Raporlama İşlemleri":
    rapor_turu = st.sidebar.selectbox(
        "📄 Rapor Türü Seçin:",
        ["-- Seçiniz --", "♻️ AYP (Atık Yönetim Planı) Raporu"],
    )
    if rapor_turu == "♻️ AYP (Atık Yönetim Planı) Raporu":
        render_ayp_module()

elif islem_kategorisi == "🏗️ Yıkım Planı ve Yasal Evrak Modülü":
    pass

elif islem_kategorisi == "🧪 ISO/IEC 17025 Kalite Yönetimi":
    render_kalite_yonetim_module()

if islem_kategorisi == "-- Seçiniz --":
    st.markdown("### 🏢 Asbest ve Atık Yönetim Rapor Sistemi")
    st.info(
        "💡 Lütfen sol menüden yapacağınız işlem kategorisini seçerek devam"
        " edin."
    )

elif islem_kategorisi == "📊 Raporlama İşlemleri":
    pass

elif islem_kategorisi == "🏗️ Yıkım Planı ve Yasal Evrak Modülü":
    pass

elif islem_kategorisi == "🧪 ISO/IEC 17025 Kalite Yönetimi":
    render_kalite_yonetim_module()

# ---------------------------------------------------------
# 1. KATEGORİ: RAPORLAMA İŞLEMLERİ (Asbest, Toz, AYP)
# ---------------------------------------------------------
if ana_kategori == "📊 Raporlama İşlemleri":
    rapor_turu = st.sidebar.selectbox(
        "📝 Rapor Türü Seçin:",
        [
            "-- Seçiniz --",
            "🔬 Asbest Tür Tayini Raporu",
            "💨 Toz Ölçüm Raporu",
            "♻️ AYP (Atık Yönetim Planı) Raporu"
        ]
    )
    
    if rapor_turu == "-- Seçiniz --":
        st.title("📊 Raporlama İşlemleri")
        st.warning("⚠️ Lütfen sol menüden oluşturmak istediğiniz **Rapor Türünü** seçin.")
    elif rapor_turu == "🔬 Asbest Tür Tayini Raporu":
        render_asbest_module()
    elif rapor_turu == "💨 Toz Ölçüm Raporu":
        render_toz_module()
    elif rapor_turu == "♻️ AYP (Atık Yönetim Planı) Raporu":
        render_ayp_module()

# ---------------------------------------------------------
# 2. KATEGORİ: YIKIM PLAN VE YASAL EVRAK MODÜLÜ
# ---------------------------------------------------------
elif ana_kategori == "🏗️ Yıkım Planı ve Yasal Evrak Modülü":
    render_yikim_module()

# ---------------------------------------------------------
# SEÇİM YAPILMADIĞINDA GÖSTERİLECEK KARŞILAMA EKRANI
# ---------------------------------------------------------
else:
    st.title("🏢 Asbest ve Atık Yönetim Rapor Sistemi")
    st.info("💡 Lütfen sol menüden yapacağınız işlem kategorisini seçerek devam edin.")

def parse_asbest_tutanak(file):
    df_raw = pd.read_excel(file, header=None)
    
    info = {
        'musteri_adi': '',
        'adres': '-',
        'pafta': '-',
        'ada': '-',
        'parsel': '-',
        'numune_tarihi': datetime.now().strftime("%d.%m.%Y"),
        'teklif_no': '',
        'telefon': '-'
    }
    
    # 1. Üst Bilgileri Okuma
    for idx in range(min(25, len(df_raw))):
        row_values = [str(x).strip() for x in df_raw.iloc[idx].values if pd.notna(x)]
        row_text = " ".join(row_values)
        
        if "Talep Numarası" in row_text or "Teklif" in row_text:
            for cell in row_values:
                m = re.search(r'(\d{2}-\d{2}-\d+)', cell)
                if m:
                    info['teklif_no'] = m.group(1)
        
        if "Firma Adı:" in row_text:
            m = re.search(r'Firma Adı:\s*(.*?)(?:Telefon|Adres|$)', row_text, re.IGNORECASE)
            if m and m.group(1).strip():
                info['musteri_adi'] = m.group(1).strip()
        
        if "Telefon Numarası:" in row_text:
            m = re.search(r'Telefon Numarası:\s*(.*)', row_text, re.IGNORECASE)
            if m and m.group(1).strip():
                info['telefon'] = m.group(1).strip()

        if "Firma Adresi:" in row_text:
            m = re.search(r'Firma Adresi:\s*(.*)', row_text, re.IGNORECASE)
            if m and m.group(1).strip():
                info['adres'] = m.group(1).strip()
                
        if "Pafta No:" in row_text or "Parsel No:" in row_text:
            p = re.search(r'Pafta\s*No:\s*([^\s|]*)(?=\s*Ada|$)', row_text, re.IGNORECASE)
            a = re.search(r'Ada\s*No:\s*([^\s|]*)(?=\s*Parsel|$)', row_text, re.IGNORECASE)
            pr = re.search(r'Parsel\s*No:\s*([^\s|]*)(?=$)', row_text, re.IGNORECASE)
            
            if p and p.group(1).strip(): info['pafta'] = p.group(1).strip()
            if a and a.group(1).strip(): info['ada'] = a.group(1).strip()
            if pr and pr.group(1).strip(): info['parsel'] = pr.group(1).strip()

        if "Tarih" in row_text:
            for cell in row_values:
                if re.match(r'\d{2}\.\d{2}\.\d{4}', cell):
                    info['numune_tarihi'] = cell

    # 2. Hata Ayıklama (Debug) İçin Yakalanan Tüm Satırları Toplama
    samples = []
    seen_codes = set()
    debug_detected_rows = []

    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        non_empty = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != '']
        row_str = " ".join(non_empty)
        
        # Eğer satırda NK kelimesi veya numune formatı geçiyorsa yakala
        if "NK" in row_str:
            debug_detected_rows.append(f"Satır {idx}: {non_empty}")
            
            code_match = re.search(r'(NK[^\s,]*)', row_str)
            if code_match:
                code = code_match.group(1)
                if code not in seen_codes:
                    seen_codes.add(code)
                    
                    code_idx = -1
                    for i, val in enumerate(non_empty):
                        if code in val:
                            code_idx = i
                            break
                    
                    tur = non_empty[code_idx + 1] if len(non_empty) > code_idx + 1 else "-"
                    yer = non_empty[code_idx + 2] if len(non_empty) > code_idx + 2 else "-"
                    yontem = non_empty[code_idx + 3] if len(non_empty) > code_idx + 3 else "TS EN ISO 16000-7"
                    strateji = non_empty[code_idx + 4] if len(non_empty) > code_idx + 4 else "Görsel ve Alansal"

                    samples.append({
                        'kod': code,
                        'tur': tur,
                        'yer': yer,
                        'yontem': yontem,
                        'strateji': strateji
                    })

    # Ekranın üst kısmına yakalanan tüm ham satırları basıyoruz
    st.expander("🔍 Hata Ayıklama: Excel'den Okunan Tüm NK Satırları", expanded=True).write(debug_detected_rows)

    return info, samples

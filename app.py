import streamlit as st
from modules.asbest import render_asbest_module
from modules.toz import render_toz_module
from modules.ayp import render_ayp_module
from modules.yikim_plani_modulu import render as render_yikim_plani

st.set_page_config(
    page_title="Asya Asbest & Atık Yönetim Sistemi",
    page_icon="🔬",
    layout="wide"
)

# --- Yan Menü (Sidebar) ---
with st.sidebar:
    st.image("https://img.icons8.com/color/96/experimental-copy.png", width=80)
    st.markdown("### 🔬 Laboratuvar Modülü")
    st.write("ASYA Asbest Danışmanlık ve Laboratuvar Hizmetleri Otomasyon Paneli")
    st.markdown("---")
    rapor_turu = st.selectbox(
        "📋 İşlem / Rapor Türü Seçin:",
        [
            "-- Seçiniz --",
            "🔬 Asbest Tür Tayini Raporu",
            "💨 Toz Raporu",
            "♻️ AYP (Atık Yönetim Planı) Raporu",
            "🏗️ Yıkım Planı ve Yasal Evrak Modülü"
        ]
    )
    st.markdown("---")

# --- Ana Ekran Yönlendirmeleri ---
st.title("🏢 Asbest ve Atık Yönetim Rapor Sistemi")
st.markdown("---")

if rapor_turu == "-- Seçiniz --":
    st.warning("⚠️ Lütfen sol menüden oluşturmak istediğiniz **Rapor Türünü** seçin.")
elif rapor_turu == "🔬 Asbest Tür Tayini Raporu":
    render_asbest_module()
elif rapor_turu == "💨 Toz Raporu":
    render_toz_module()
elif rapor_turu == "♻️ AYP (Atık Yönetim Planı) Raporu":
    render_ayp_module()
elif rapor_turu == "🏗️ Yıkım Planı ve Yasal Evrak Modülü":
    render_yikim_plani()

# modules/asbest.py içerisindeki parse_asbest_tutanak fonksiyonunu bu şekilde güncelleyin:

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

    # 2. Numune Tablosunu Okuma (Mükerrer Tespiti & Dinamik İndeksleme)
    samples = []
    seen_codes = set() # 15 numune çıkmasını engelleyen küme

    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        non_empty = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != '']
        row_str = " ".join(non_empty)
        
        code_match = re.search(r'\b(NK\.\d{2}\.\d+-\d+)\b', row_str)
        if code_match:
            code = code_match.group(1)
            
            # Alt tablolardaki yinelenen kodları süzme
            if code not in seen_codes:
                seen_codes.add(code)
                
                # Kodun konumunu dinamik bulup sonraki hücreleri sırayla eşleme
                code_idx = -1
                for i, val in enumerate(non_empty):
                    if code in val:
                        code_idx = i
                        break
                
                tur = non_empty[code_idx + 1] if len(non_empty) > code_idx + 1 else "Beton / Sıva"
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

    return info, samples

import os
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

st.set_page_config(
    page_title="Asya Asbest & Atık Yönetim Sistemi",
    page_icon="🧪",
    layout="wide"
)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def read_tutanak_details(tutanak_path):
    """Excel tutanağından firma, adres ve pafta/ada/parsel bilgilerini okur"""
    try:
        df = pd.read_excel(tutanak_path, sheet_name='Table 1', header=None)
    except:
        xls = pd.ExcelFile(tutanak_path)
        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)

    teklif_no = str(df.iloc[3, 0]).strip() if pd.notna(df.iloc[3, 0]) else "-"
    numune_tarihi = str(df.iloc[3, 5]).split()[0] if pd.notna(df.iloc[3, 5]) else "-"
    
    raw_firma = str(df.iloc[4, 0]) if pd.notna(df.iloc[4, 0]) else ""
    musteri_adi = raw_firma.replace("Firma Adı:", "").strip()

    raw_adres = str(df.iloc[5, 0]) if pd.notna(df.iloc[5, 0]) else ""
    adres = raw_adres.replace("Firma Adresi:", "").strip()

    raw_pafta = str(df.iloc[6, 0]) if pd.notna(df.iloc[6, 0]) else "-"
    raw_ada = str(df.iloc[6, 4]) if pd.notna(df.iloc[6, 4]) else "-"
    raw_parsel = str(df.iloc[6, 8]) if pd.notna(df.iloc[6, 8]) else "-"

    pafta = raw_pafta.replace("Pafta No:", "").strip() or "-"
    ada = raw_ada.replace("Ada No:", "").strip() or "-"
    parsel = raw_parsel.replace("Parsel No:", "").strip() or "-"
    pafta_ada_parsel = f"{pafta} / {ada} / {parsel}"

    context = {
        'musteri_adi': musteri_adi,
        'MUSTERI_ADI': musteri_adi,
        'firma_adi': musteri_adi,
        'FIRMA_ADI': musteri_adi,
        
        'adres': adres,
        'ADRES': adres,
        'santiye_adresi': adres,
        'SANTIYE_ADRESI': adres,
        
        'teklif_no': teklif_no,
        'numune_tarihi': numune_tarihi,
        'pafta': pafta,
        'ada': ada,
        'parsel': parsel,
        'pafta_ada_parsel': pafta_ada_parsel,
        'PAFTA_ADA_PARSEL': pafta_ada_parsel
    }
    return context

# --- Yan Menü (Sidebar) Tasarımı ---
with st.sidebar:
    st.image("https://img.icons8.com/color/96/experimental-copy.png", width=80)
    st.markdown("### 🔬 Laboratuvar Modülü")
    st.write("ASYA Asbest Danışmanlık ve Laboratuvar Hizmetleri Otomasyon Paneli")
    st.markdown("---")
    rapor_turu = st.selectbox(
        "📋 İşlem / Rapor Türü Seçin:", 
        [
            "-- Seçiniz --", 
            "🧪 Asbest Tür Tayini Raporu", 
            "💨 Toz Raporu", 
            "♻️ AYP (Atık Yönetim Planı) Raporu"
        ]
    )
    st.markdown("---")

# --- Ana Ekran Tasarımı ---
st.title("🧪 Asbest ve Atık Yönetim Rapor Sistemi")
st.markdown("---")

if rapor_turu == "-- Seçiniz --":
    st.warning("⚠️ Lütfen sol menüden oluşturmak istediğiniz **Rapor Türünü** seçin.")

elif rapor_turu == "💨 Toz Raporu":
    st.subheader("💨 Toz Ölçüm Raporu Oluşturucu")
    tutanak_file = st.file_uploader("📂 Tutanak Dosyası (Excel):", type=["xlsx", "xls"], key="toz_tutanak")
    
    if tutanak_file:
        try:
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())

            info = read_tutanak_details(tutanak_path)
            st.success("✅ Toz tutanak dosyası başarıyla okundu.")

            if st.button("🚀 Toz Raporunu Oluştur ve İndir", type="primary"):
                if os.path.exists('sablon_toz.docx'):
                    doc = DocxTemplate('sablon_toz.docx')
                    doc.render(info)

                    output_path = os.path.join(UPLOAD_FOLDER, 'Toz_Raporu_Cikti.docx')
                    doc.save(output_path)
                    st.success("✅ Toz Raporu başarıyla oluşturuldu!")
                    with open(output_path, "rb") as f:
                        st.download_button(
                            "📥 Toz Raporunu İndir (.docx)", 
                            f, 
                            file_name=f"Toz_Raporu_{info['musteri_adi']}.docx", 
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                else:
                    st.error("❌ Ana dizinde 'sablon_toz.docx' dosyası bulunamadı!")
        except Exception as e:
            st.error(f"❌ Toz raporu işlenirken hata oluştu: {e}")

elif rapor_turu == "🧪 Asbest Tür Tayini Raporu":
    st.subheader("🧬 Asbest Tür Tayini Raporu Oluşturucu")
    tutanak_file = st.file_uploader("📂 Numune Alma Tutanağı (Excel) Seçin:", type=['xlsx', 'xls'], key="asbest_tutanak")
    
    if tutanak_file:
        try:
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())

            info = read_tutanak_details(tutanak_path)
            teklif_kodu = info['teklif_no'].split("-")[-1] if "-" in info['teklif_no'] else "0000"
            info['rapor_no'] = f"ARK.26.{teklif_kodu}"

            df = pd.read_excel(tutanak_path, sheet_name='Table 1', header=None)
            numuneler = []
            for i in range(9, len(df)):
                sira_no = df.iloc[i, 0]
                numune_kodu = df.iloc[i, 1]
                if pd.isna(sira_no) or pd.isna(numune_kodu):
                    continue
                numuneler.append({
                    'sira': int(sira_no),
                    'tarih': info['numune_tarihi'],
                    'numune_kodu': str(numune_kodu).strip(),
                    'malzeme_turu': str(df.iloc[i, 4]).strip() if pd.notna(df.iloc[i, 4]) else "-",
                    'numune_alma_yeri': str(df.iloc[i, 7]).strip() if pd.notna(df.iloc[i, 7]) else "-",
                    'yontem': str(df.iloc[i, 8]).strip() if pd.notna(df.iloc[i, 8]) else "-",
                    'strateji': str(df.iloc[i, 9]).strip() if pd.notna(df.iloc[i, 9]) else "-",
                })

            st.success(f"✅ Tutanak başarıyla okundu! Toplam **{len(numuneler)}** numune tespit edildi.")
            
            if st.button("🚀 Asbest Raporunu Oluştur ve İndir", type="primary"):
                if os.path.exists('sablon.docx'):
                    doc = DocxTemplate('sablon.docx')
                    
                    # Asbest raporu için context verilerini ve dinamik numune listesini birleştiriyoruz
                    context = {
                        **info,
                        'numuneler': numuneler
                    }
                    
                    doc.render(context)

                    output_path = os.path.join(UPLOAD_FOLDER, 'Asbest_Raporu_Cikti.docx')
                    doc.save(output_path)
                    st.success("✅ Asbest Raporu başarıyla oluşturuldu!")
                    with open(output_path, "rb") as f:
                        st.download_button(
                            "📥 Asbest Raporunu İndir (.docx)", 
                            f, 
                            file_name=f"Asbest_Raporu_{info['musteri_adi']}.docx", 
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                else:
                    st.error("❌ Ana dizinde 'sablon.docx' dosyası bulunamadı!")
        except Exception as e:
            st.error(f"❌ İşlem sırasında hata oluştu: {e}")

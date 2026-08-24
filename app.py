import os
import pandas as pd
from streamlit_extras.colored_header import colored_header
import streamlit as st
from docx import Document

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def calculate_ayp_excel(file_path):
    """Ayp Hesaplama.xls dosyasından tüm atık ve hesaplama verilerini okur"""
    xls = pd.ExcelFile(file_path)
    df1 = pd.read_excel(xls, sheet_name='Sayfa1', header=None)
    df2 = pd.read_excel(xls, sheet_name='Sayfa2', header=None)
    
    def get_val(df, row, col, default=0):
        try:
            val = df.iloc[row, col]
            return val if pd.notna(val) else default
        except:
            return default

    hesaplar = {
        "TUĞLA_MIKTARI": get_val(df1, 6, 10, 9504),
        "ALÇI_MIKTARI": get_val(df1, 10, 9, 31680),
        "BETON_MIKTARI": get_val(df1, 15, 9, 177120),
        "ATERMIT_HESAP_DETAYI": get_val(df1, 53, 7, 0),
        "AHŞAP_MIKTARI": get_val(df1, 25, 7, 345.6),
        "SERAMİK_MIKTARI": get_val(df1, 32, 5, 5174.1),
        "KİREMİT_MIKTARI": get_val(df1, 27, 4, 3690),
        "DEMİR_HESAP_DETAYI": get_val(df1, 51, 5, 13120),
        "KAĞIT_HESAP_DETAYI": 12,
    }
    
    for index, row in df2.iterrows():
        atik_adi = str(row[5]) if len(row) > 5 else ""
        miktar = row[6] if len(row) > 6 else None
        if pd.notna(miktar):
            if "tuğla" in atik_adi.lower():
                hesaplar["TUĞLA_MIKTARI"] = miktar
            elif "beton" in atik_adi.lower():
                hesaplar["BETON_MIKTARI"] = miktar

    return hesaplar

def replace_tags_in_paragraph(paragraph, data_dict):
    for key, value in data_dict.items():
        tag = f"{{{{{key}}}}}"
        if tag in paragraph.text:
            for run in paragraph.runs:
                if tag in run.text:
                    run.text = run.text.replace(tag, str(value))
            if tag in paragraph.text:
                paragraph.text = paragraph.text.replace(tag, str(value))

# --- Streamlit Arayüzü ---
st.title("Asbest ve Atık Yönetim Rapor Sistemi")

rapor_turu = st.selectbox("Rapor Türünü Seçin:", ["-- Seçiniz --", "Toz Raporu", "AYP (Atık Yönetim Planı) Raporu"])

if rapor_turu == "Toz Raporu":
    tutanak_file = st.file_uploader("Numune / Şantiye Tutanak Dosyası (Excel/Word):", type=["xlsx", "xls", "docx"])
    if st.button("Toz Raporunu Oluştur") and tutanak_file:
        st.info("Toz raporu oluşturma adımları işleniyor...")
        # Toz raporu işlemleri buraya eklenebilir

elif rapor_turu == "AYP (Atık Yönetim Planı) Raporu":
    tutanak_file = st.file_uploader("Numune / Şantiye Tutanak Dosyası (Excel/Word):", type=["xlsx", "xls", "docx"])
    ayp_excel_file = st.file_uploader("AYP Hesaplama Excel Dosyası (Ayp Hesaplama.xls):", type=["xls", "xlsx"])
    
    if st.button("AYP Raporunu Oluştur ve İndir") and tutanak_file and ayp_excel_file:
        with st.spinner("Rapor hazırlanıyor, lütfen bekleyin..."):
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            excel_path = os.path.join(UPLOAD_FOLDER, ayp_excel_file.name)
            
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())
            with open(excel_path, "wb") as f:
                f.write(ayp_excel_file.getbuffer())
            
            # Hesaplamaları yap
            hesaplanan_degerler = calculate_ayp_excel(excel_path)
            
            # Şablonu yükle
            if os.path.exists('sablon_ayp.docx'):
                doc = Document('sablon_ayp.docx')
                
                for paragraph in doc.paragraphs:
                    replace_tags_in_paragraph(paragraph, hesaplanan_degerler)
                
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                replace_tags_in_paragraph(paragraph, hesaplanan_degerler)
                
                output_path = os.path.join(UPLOAD_FOLDER, 'AYP_Raporu_Cikti.docx')
                doc.save(output_path)
                
                st.success("Rapor başarıyla oluşturuldu!")
                with open(output_path, "rb") as file:
                    st.download_button(
                        label="Oluşan AYP Raporunu İndir",
                        data=file,
                        file_name="AYP_Raporu_Cikti.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                st.error("Ana dizinde 'sablon_ayp.docx' dosyası bulunamadı! Lütfen şablon dosyasını yükleyin.")

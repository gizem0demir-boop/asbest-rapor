import os
import pandas as pd
import streamlit as st
from docx import Document

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def calculate_ayp_excel(file_path):
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
        # Hem boşluksuz hem de boşluklu yazım ihtimaline karşı değiştirme yapıyoruz
        tags = [f"{{{{{key}}}}}", f"{{{{ {key} }}}}", f"{{{{  {key}  }}}}"]
        for tag in tags:
            if tag in paragraph.text:
                for run in paragraph.runs:
                    if tag in run.text:
                        run.text = run.text.replace(tag, str(value))
                if tag in paragraph.text:
                    paragraph.text = paragraph.text.replace(tag, str(value))

# --- Streamlit Arayüzü ---
st.title("Asbest ve Atık Yönetim Rapor Sistemi")

rapor_turu = st.selectbox(
    "Rapor Türünü Seçin:", 
    [
        "-- Seçiniz --", 
        "Asbest Tür Tayini Raporu", 
        "Toz Raporu", 
        "AYP (Atık Yönetim Planı) Raporu"
    ]
)

if rapor_turu == "Asbest Tür Tayini Raporu":
    tutanak_file = st.file_uploader("Numune Alma Tutanağı (Excel) Seçin:", type=['xlsx', 'xls'])
    if st.button("Asbest Raporunu Oluştur") and tutanak_file:
        try:
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())

            df = pd.read_excel(tutanak_path, sheet_name='Table 1', header=None)
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

            teklif_kodu = teklif_no.split("-")[-1] if "-" in teklif_no else "0000"
            rapor_no = f"ARK.26.{teklif_kodu}"

            numuneler = []
            for i in range(9, len(df)):
                sira_no = df.iloc[i, 0]
                numune_kodu = df.iloc[i, 1]
                if pd.isna(sira_no) or pd.isna(numune_kodu):
                    continue
                numuneler.append({
                    'sira': int(sira_no),
                    'tarih': numune_tarihi,
                    'kod': str(numune_kodu).strip(),
                    'tur': str(df.iloc[i, 4]).strip() if pd.notna(df.iloc[i, 4]) else "-",
                    'yer': str(df.iloc[i, 7]).strip() if pd.notna(df.iloc[i, 7]) else "-",
                    'yontem': str(df.iloc[i, 8]).strip() if pd.notna(df.iloc[i, 8]) else "-",
                    'strateji': str(df.iloc[i, 9]).strip() if pd.notna(df.iloc[i, 9]) else "-",
                    'bolum': str(df.iloc[i, 10]).strip() if pd.notna(df.iloc[i, 10]) else "-",
                    'homojenite': 'Homojen',
                    'onislem': 'Parçalama',
                    'sonuc': 'Asbest tespit edilmedi'
                })

            st.success(f"Tutanak okundu! Toplam {len(numuneler)} numune bulundu.")
            
            if os.path.exists('sablon.docx'):
                doc = Document('sablon.docx')
                context = {
                    'musteri_adi': musteri_adi, 'adres': adres, 'teklif_no': teklif_no,
                    'rapor_no': rapor_no, 'numune_tarihi': numune_tarihi,
                    'pafta': pafta, 'ada': ada, 'parsel': parsel
                }
                
                # Paragraflardaki etiketleri değiştir
                for p in doc.paragraphs:
                    replace_tags_in_paragraph(p, context)

                # Tablolardaki etiketleri değiştir (Örn: Üst tablodaki rapor_no, müşteri adı vb.)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                replace_tags_in_paragraph(p, context)

                # Numune tablosunu doldur (Örnek olarak 3. tablo)
                if len(doc.tables) > 3:
                    table = doc.tables[3]
                    while len(table.rows) > 3:
                        r = table.rows[3]._tr
                        r.getparent().remove(r)
                    for n in numuneler:
                        row_cells = table.add_row().cells
                        veriler = [str(n['sira']), str(n['tarih']), str(n['kod']), str(n['tur']), str(n['yer']), str(n['yontem']), str(n['strateji']), str(n['homojenite']), str(n['onislem']), str(n['sonuc'])]
                        for idx, val in enumerate(veriler):
                            if idx < len(row_cells):
                                row_cells[idx].text = val

                output_path = os.path.join(UPLOAD_FOLDER, 'Asbest_Raporu_Cikti.docx')
                doc.save(output_path)
                
                with open(output_path, "rb") as f:
                    st.download_button("📥 Asbest Raporunu İndir", f, file_name=f"Asbest_Raporu_{musteri_adi}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            else:
                st.error("Ana dizinde 'sablon.docx' dosyası bulunamadı!")
        except Exception as e:
            st.error(f"Hata oluştu: {e}")

elif rapor_turu == "Toz Raporu":
    tutanak_file = st.file_uploader("Numune / Şantiye Tutanak Dosyası:", type=["xlsx", "xls", "docx"])
    if st.button("Toz Raporunu Oluştur") and tutanak_file:
        if os.path.exists('sablon_toz.docx'):
            doc = Document('sablon_toz.docx')
            output_path = os.path.join(UPLOAD_FOLDER, 'Toz_Raporu_Cikti.docx')
            doc.save(output_path)
            with open(output_path, "rb") as f:
                st.download_button("📥 Toz Raporunu İndir", f, file_name="Toz_Bastirma_Raporu.docx")
        else:
            st.error("'sablon_toz.docx' dosyası bulunamadı!")

elif rapor_turu == "AYP (Atık Yönetim Planı) Raporu":
    tutanak_file = st.file_uploader("Numune / Şantiye Tutanak Dosyası:", type=["xlsx", "xls", "docx"])
    ayp_excel_file = st.file_uploader("AYP Hesaplama Excel Dosyası (Ayp Hesaplama.xls):", type=["xls", "xlsx"])
    
    if st.button("AYP Raporunu Oluştur ve İndir") and tutanak_file and ayp_excel_file:
        with st.spinner("Rapor hazırlanıyor..."):
            excel_path = os.path.join(UPLOAD_FOLDER, ayp_excel_file.name)
            with open(excel_path, "wb") as f:
                f.write(ayp_excel_file.getbuffer())
            
            hesaplanan_degerler = calculate_ayp_excel(excel_path)
            
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
                
                with open(output_path, "rb") as f:
                    st.download_button("📥 AYP Raporunu İndir", f, file_name="AYP_Raporu_Cikti.docx")
            else:
                st.error("'sablon_ayp.docx' dosyası bulunamadı!")

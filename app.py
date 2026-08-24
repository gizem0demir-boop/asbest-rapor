import os
import pandas as pd
import streamlit as st
from docx import Document

# --- Sayfa Yapılandırması ---
st.set_page_config(
    page_title="Asya Asbest & Atık Yönetim Sistemi",
    page_icon="🧪",
    layout="wide"
)

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
        tags = [f"{{{{{key}}}}}", f"{{{{ {key} }}}}", f"{{{{  {key}  }}}}"]
        for tag in tags:
            if tag in paragraph.text:
                for run in paragraph.runs:
                    if tag in run.text:
                        run.text = run.text.replace(tag, str(value))
                if tag in paragraph.text:
                    paragraph.text = paragraph.text.replace(tag, str(value))

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
    st.info("💡 İpucu: Excel tutanak dosyalarınızı eksiksiz yüklediğinizden emin olun.")

# --- Ana Ekran Tasarımı ---
st.title("🧪 Asbest ve Atık Yönetim Rapor Sistemi")
st.markdown("Laboratuvar analiz verilerinizi ve şablonlarınızı hızlıca rapora dönüştürün.")
st.markdown("---")

if rapor_turu == "-- Seçiniz --":
    st.warning("⚠️ Lütfen sol menüden oluşturmak istediğiniz **Rapor Türünü** seçin.")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("### 🧬 Asbest Analizi")
        st.write("Numune tutanaklarından otomatik asbest tür tayini raporu üretin.")
    with col2:
        st.markdown("### 🥽 Toz Ölçümleri")
        st.write("Şantiye ortam toz ölçüm ve bastırma raporlarını hazırlayın.")
    with col3:
        st.markdown("### 📊 Atık Yönetimi")
        st.write("Excel hesaplamalarını entegre ederek AYP raporunuzu oluşturun.")

elif rapor_turu == "🧪 Asbest Tür Tayini Raporu":
    st.subheader("🧬 Asbest Tür Tayini Raporu Oluşturucu")
    tutanak_file = st.file_uploader("📂 Numune Alma Tutanağı (Excel) Seçin:", type=['xlsx', 'xls'])
    
    if tutanak_file:
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
                    'Sıra': int(sira_no),
                    'Tarih': numune_tarihi,
                    'Numune Kodu': str(numune_kodu).strip(),
                    'Malzeme Türü': str(df.iloc[i, 4]).strip() if pd.notna(df.iloc[i, 4]) else "-",
                    'Numune Alma Yeri': str(df.iloc[i, 7]).strip() if pd.notna(df.iloc[i, 7]) else "-",
                    'Yöntem': str(df.iloc[i, 8]).strip() if pd.notna(df.iloc[i, 8]) else "-",
                    'Strateji': str(df.iloc[i, 9]).strip() if pd.notna(df.iloc[i, 9]) else "-",
                    'Bölüm': str(df.iloc[i, 10]).strip() if pd.notna(df.iloc[i, 10]) else "-",
                })

            st.success(f"✅ Tutanak başarıyla okundu! Toplam **{len(numuneler)}** numune tespit edildi.")
            
            # --- DOSYA YÜKLENİR YÜKLENMEZ AÇILAN ÖN İZLEME KUTUSU (EXPANDER) ---
            with st.expander("🔍 Tutanağından Okunan Bilgileri ve Numune Listesini İncele", expanded=True):
                col_a, col_b = st.columns(2)
                with col_a:
                    st.markdown(f"**Müşteri Adı:** {musteri_adi}")
                    st.markdown(f"**Teklif Numarası:** {teklif_no}")
                    st.markdown(f"**Rapor Numarası:** {rapor_no}")
                with col_b:
                    st.markdown(f"**Numune Tarihi:** {numune_tarihi}")
                    st.markdown(f"**Pafta / Ada / Parsel:** {pafta} / {ada} / {parsel}")
                
                st.markdown(f"**Adres:** {adres}")
                st.markdown("---")
                st.markdown("### Okunan Numune Listesi")
                df_preview = pd.DataFrame(numuneler)
                st.dataframe(df_preview, use_container_width=True)

            # --- Word Raporu Oluşturma Butonu ---
            if st.button("🚀 Asbest Raporunu Word Olarak Oluştur", type="primary"):
                if os.path.exists('sablon.docx'):
                    doc = Document('sablon.docx')
                    context = {
                        'musteri_adi': musteri_adi, 'adres': adres, 'teklif_no': teklif_no,
                        'rapor_no': rapor_no, 'numune_tarihi': numune_tarihi,
                        'pafta': pafta, 'ada': ada, 'parsel': parsel
                    }
                    
                    for p in doc.paragraphs:
                        replace_tags_in_paragraph(p, context)

                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    replace_tags_in_paragraph(p, context)

                    if len(doc.tables) > 3:
                        table = doc.tables[3]
                        while len(table.rows) > 3:
                            r = table.rows[3]._tr
                            r.getparent().remove(r)
                        for n in numuneler:
                            row_cells = table.add_row().cells
                            veriler = [str(n['Sıra']), str(n['Tarih']), str(n['Numune Kodu']), str(n['Malzeme Türü']), str(n['Numune Alma Yeri']), str(n['Yöntem']), str(n['Strateji']), 'Homojen', 'Parçalama', 'Asbest tespit edilmedi']
                            for idx, val in enumerate(veriler):
                                if idx < len(row_cells):
                                    row_cells[idx].text = val

                    output_path = os.path.join(UPLOAD_FOLDER, 'Asbest_Raporu_Cikti.docx')
                    doc.save(output_path)
                    
                    st.success("✅ Asbest Raporu başarıyla oluşturuldu!")
                    with open(output_path, "rb") as f:
                        st.download_button("📥 Asbest Raporunu İndir (.docx)", f, file_name=f"Asbest_Raporu_{musteri_adi}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                else:
                    st.error("❌ Ana dizinde 'sablon.docx' dosyası bulunamadı!")
                    
        except Exception as e:
            st.error(f"❌ İşlem sırasında hata oluştu: {e}")

elif rapor_turu == "💨 Toz Raporu":
    st.subheader("💨 Toz Ölçüm Raporu Oluşturucu")
    tutanak_file = st.file_uploader("📂 Numune / Şantiye Tutanak Dosyası:", type=["xlsx", "xls", "docx"])
    if st.button("🚀 Toz Raporunu Oluştur", type="primary") and tutanak_file:
        if os.path.exists('sablon_toz.docx'):
            doc = Document('sablon_toz.docx')
            output_path = os.path.join(UPLOAD_FOLDER, 'Toz_Raporu_Cikti.docx')
            doc.save(output_path)
            st.success("✅ Toz raporu hazırlandı!")
            with open(output_path, "rb") as f:
                st.download_button("📥 Toz Raporunu İndir (.docx)", f, file_name="Toz_Bastirma_Raporu.docx")
        else:
            st.error("❌ Ana dizinde 'sablon_toz.docx' dosyası bulunamadı!")

elif rapor_turu == "♻️ AYP (Atık Yönetim Planı) Raporu":
    st.subheader("♻️ Atık Yönetim Planı (AYP) Raporu Oluşturucu")
    tutanak_file = st.file_uploader("📂 Numune / Şantiye Tutanak Dosyası:", type=["xlsx", "xls", "docx"])
    ayp_excel_file = st.file_uploader("📊 AYP Hesaplama Excel Dosyası (Ayp Hesaplama.xls):", type=["xls", "xlsx"])
    
    if st.button("🚀 AYP Raporunu Oluştur ve İndir", type="primary") and tutanak_file and ayp_excel_file:
        with st.spinner("⏳ Rapor hesaplanıyor ve şablona işleniyor..."):
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
                
                st.success("✅ AYP Raporu başarıyla oluşturuldu!")
                with open(output_path, "rb") as f:
                    st.download_button("📥 AYP Raporunu İndir (.docx)", f, file_name="AYP_Raporu_Cikti.docx")
            else:
                st.error("❌ Ana dizinde 'sablon_ayp.docx' dosyası bulunamadı!")

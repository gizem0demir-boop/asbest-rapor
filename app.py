import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os

st.set_page_config(page_title="Asbest & Toz Rapor Otomasyonu", page_icon="🧪", layout="centered")

st.title("🧪 Rapor Oluşturma Otomasyonu")
st.markdown("Excel tutanağınızı yükleyin, istediğiniz raporu saniyeler içinde hazırlayın.")

# Rapor türü seçimi için yan yana butonlar veya radyo düğmesi
rapor_tipi = st.radio(
    "Oluşturulacak Rapor Türünü Seçin:",
    ["Asbest Tür Tayini Raporu", "Toz Bastırma Raporu"]
)

uploaded_file = st.file_uploader("Numune Alma Tutanağı (Excel) Seçin", type=['xlsx'])

if uploaded_file is not None:
    try:
        # Excel sayfasını okuma
        df = pd.read_excel(uploaded_file, sheet_name='Table 1', header=None)

        # Üst Bilgileri Ayrıştırma
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

        # Otomatik Rapor Numarası Üretme (ARK.26.XXXX)
        teklif_kodu = teklif_no.split("-")[-1] if "-" in teklif_no else "0000"
        rapor_no = f"ARK.26.{teklif_kodu}"

        # Numune Listesini Çıkarma (9. satırdan itibaren)
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

        st.success(f"✅ Tutanak başarıyla okundu! Toplam **{len(numuneler)}** adet numune tespit edildi.")
        
        # Önizleme Gösterimi
        with st.expander("📌 Okunan Müşteri & Tutanak Bilgilerini Kontrol Et"):
            st.write(f"**Seçilen Rapor:** {rapor_tipi}")
            st.write(f"**Müşteri Adı:** {musteri_adi}")
            st.write(f"**Adres:** {adres}")
            st.write(f"**Teklif No / Rapor No:** {teklif_no} / {rapor_no}")
            st.write(f"**Pafta/Ada/Parsel:** {pafta} / {ada} / {parsel}")
            st.write(f"**Numune Tarihi:** {numune_tarihi}")

        # Rapor Oluşturma Butonu
        if st.button("🚀 Word Raporunu Oluştur"):
            
            # Hangi rapor seçildiyse ona uygun şablonu belirliyoruz
            if rapor_tipi == "Toz Bastırma Raporu":
                sablon_adi = "sablon_toz.docx"
                dosya_oneki = "Toz_Bastirma_Raporu"
            else:
                sablon_adi = "sablon.docx"
                dosya_oneki = "Asbest_Raporu"

            if not os.path.exists(sablon_adi):
                st.error(f"❌ '{sablon_adi}' dosyası bulutta bulunamadı! Lütfen ilgili Word şablonunu GitHub deposuna yükleyin.")
            else:
                doc = DocxTemplate(sablon_adi)
                context = {
                    'musteri_adi': musteri_adi,
                    'adres': adres,
                    'teklif_no': teklif_no,
                    'rapor_no': rapor_no,
                    'numune_tarihi': numune_tarihi,
                    'pafta': pafta,
                    'ada': ada,
                    'parsel': parsel,
                    'numuneler': numuneler
                }
                
                doc.render(context)

                # Eğer şablonda tablo varsa (Asbest raporundaki gibi) güvenle işleme
                if len(doc.tables) > 3 and rapor_tipi == "Asbest Tür Tayini Raporu":
                    table = doc.tables[3]
                    while len(table.rows) > 3:
                        r = table.rows[3]._tr
                        r.getparent().remove(r)

                    for n in numuneler:
                        row_cells = table.add_row().cells
                        veriler = [
                            str(n.get('sira', '')),
                            str(n.get('tarih', '')),
                            str(n.get('kod', '')),
                            str(n.get('tur', '')),
                            str(n.get('yer', '')),
                            str(n.get('yontem', '')),
                            str(n.get('strateji', '')),
                            str(n.get('homojenite', '')),
                            str(n.get('onislem', '')),
                            str(n.get('sonuc', ''))
                        ]
                        for i, val in enumerate(veriler):
                            if i < len(row_cells):
                                row_cells[i].text = val

                target_stream = io.BytesIO()
                doc.save(target_stream)
                target_stream.seek(0)
                
                dosya_adi = f"{dosya_oneki}_{musteri_adi.replace(' ', '_')}.docx"
                
                st.download_button(
                    label="📥 Hazır Word Raporunu İndir",
                    data=target_stream,
                    file_name=dosya_adi,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

    except Exception as e:
        st.error(f"Excel işlenirken beklenmeyen bir hata oluştu: {e}")

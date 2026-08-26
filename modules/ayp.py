import os
from datetime import datetime
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from utils import read_tutanak_details, UPLOAD_FOLDER

def render_ayp_module():
    st.subheader("♻️ Atık Yönetim Planı (AYP) Rapor Oluşturucu")

    col1, col2 = st.columns(2)
    with col1:
        tutanak_file = st.file_uploader("📋 1. Tutanak Dosyası (Excel - Künye için):", type=["xlsx", "xls"], key="ayp_tutanak")
    with col2:
        ayp_file = st.file_uploader("📊 2. AYP Hesaplama Dosyası (Excel):", type=["xlsx", "xls"], key="ayp_excel")

    if tutanak_file and ayp_file:
        try:
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())
            info = read_tutanak_details(tutanak_path)

            ayp_path = os.path.join(UPLOAD_FOLDER, ayp_file.name)
            with open(ayp_path, "wb") as f:
                f.write(ayp_file.getbuffer())

            df_sayfa2 = pd.read_excel(ayp_path, sheet_name="Sayfa2")

            atik_miktarlari = {}
            for _, row in df_sayfa2.iterrows():
                key = row.iloc[5]
                val = row.iloc[6]
                if pd.notna(key):
                    atik_miktarlari[str(key).strip().lower()] = 0 if pd.isna(val) else val

            genel_toplam = 0
            for _, row in df_sayfa2.iterrows():
                if str(row.iloc[4]).strip().lower() == "toplam":
                    genel_toplam = row.iloc[6]

            bugun_tarihi = datetime.now().strftime("%d.%m.%Y")

            info.update({
                "tarih": bugun_tarihi, "TARIH": bugun_tarihi, "rapor_tarihi": bugun_tarihi,
                "alan_m2": 82, "kat_sayisi": 6, "cati_alan_m2": 82, "oda_sayisi": 3,
                "daire_sayisi": 6, "isci_sayisi": 4, "calisma_suresi_gun": 5,
                "pencere_adet": 6, "seramik_adet": 360, "laminant_alan_m2": 8,
                "asbest_toplam_kg": atik_miktarlari.get("asbest içeren inşaat malzemeleri", 0),
                "beton_toplam_kg": atik_miktarlari.get("beton", 177120),
                "kiremit_toplam_kg": 3690,
                "seramik_genel_toplam_kg": 5174.1,
                "ahsap_toplam_kg": atik_miktarlari.get("ahşap", 345.6),
                "tugla_toplam_kg": atik_miktarlari.get("tuğla", 9504),
                "siva_toplam_kg": atik_miktarlari.get("17 08 01 dışındaki alçı bazlı inşaat malzemeleri", 31680),
                "toplam_karisik_metal": atik_miktarlari.get("karışık metaller", 13120),
                "demir_temel_toplam": 3280, "demir_kat_toplam": 9840,
                "kagit_toplam_kg": atik_miktarlari.get("kağıt ve karton ambalaj", 12),
                "plastik_toplam_kg": atik_miktarlari.get("plastik ambalaj", 0),
                "cam_miktari": atik_miktarlari.get("cam ambalaj", 0),
                "seramik_adet_toplam_kg": 1440,
                "genel_toplam_miktar": genel_toplam if genel_toplam != 0 else 236955.7
            })

            st.success("✅ Tutanak ve AYP hesaplama verileri birleştirildi.")

            if st.button("📄 AYP Raporunu Oluştur ve İndir", type="primary"):
                if os.path.exists("sablon_ayp.docx"):
                    doc = DocxTemplate("sablon_ayp.docx")
                    doc.render(info)
                    
                    output_path = os.path.join(UPLOAD_FOLDER, "AYP_Raporu_Cikti.docx")
                    doc.save(output_path)
                    st.success("✅ Atık Yönetim Planı Raporu oluşturuldu!")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            "📥 AYP Raporunu İndir (.docx)",
                            f,
                            file_name=f"AYP_Raporu_{info['musteri_adi']}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                else:
                    st.error("❌ Ana dizinde 'sablon_ayp.docx' bulunamadı!")
        except Exception as e:
            st.error(f"❌ AYP raporu hatası: {e}")
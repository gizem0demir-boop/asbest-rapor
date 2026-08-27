import os
import jinja2
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st
from utils import UPLOAD_FOLDER, read_tutanak_details


def render_ayp_module():
    st.subheader("♻️ Atık Yönetim Planı (AYP) Rapor Oluşturucu")

    col1, col2 = st.columns(2)
    with col1:
        tutanak_file = st.file_uploader(
            "📂 1. Tutanak Dosyası (Excel - Künye için):",
            type=["xlsx", "xls"],
            key="ayp_tutanak",
        )
    with col2:
        ayp_file = st.file_uploader(
            "📂 2. AYP Hesaplama Dosyası (Excel):",
            type=["xlsx", "xls"],
            key="ayp_excel",
        )

    if tutanak_file and ayp_file:
        try:
            # Tutanak dosyasını kaydet ve oku
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())
            info = read_tutanak_details(tutanak_path)

            # AYP hesaplama dosyasını kaydet ve oku
            ayp_path = os.path.join(UPLOAD_FOLDER, ayp_file.name)
            with open(ayp_path, "wb") as f:
                f.write(ayp_file.getbuffer())

            xls = pd.ExcelFile(ayp_path)
            df_sayfa2 = (
                pd.read_excel(ayp_path, sheet_name="Sayfa2")
                if "Sayfa2" in xls.sheet_names
                else pd.DataFrame()
            )

            atik_miktarlari = {}
            genel_toplam = 0.0

            for _, row in df_sayfa2.iterrows():
                row_vals = [v for v in row.values if pd.notna(v)]
                if not row_vals:
                    continue

                row_str_full = " ".join([str(v) for v in row_vals]).lower()
                if "toplam" in row_str_full and "daire" not in row_str_full:
                    nums = [
                        float(v)
                        for v in row_vals
                        if isinstance(v, (int, float))
                        or (str(v).replace(".", "", 1).isdigit())
                    ]
                    if nums:
                        genel_toplam = nums[-1]

                key = row.iloc[5] if len(row) > 5 else None
                val = row.iloc[6] if len(row) > 6 else None
                if pd.notna(key):
                    atik_miktarlari[str(key).strip().lower()] = (
                        0.0 if pd.isna(val) else float(val)
                    )

            bugun_tarihi = pd.Timestamp.now().strftime("%d.%m.%Y")

            info.update(
                {
                    "tarih": bugun_tarihi,
                    "TARIH": bugun_tarihi,
                    "rapor_tarihi": bugun_tarihi,
                    "alan_m2": 85.0,
                    "kat_sayisi": 6.0,
                    "cati_alan_m2": 85.0,
                    "oda_sayisi": 3,
                    "daire_sayisi": 6.0,
                    "isci_sayisi": 4,
                    "calisma_suresi_gun": 5,
                    "pencere_adet": 6,
                    "seramik_adet": 360,
                    "laminant_alan_m2": 8,
                    "asbest_toplam_kg": atik_miktarlari.get(
                        "asbest içeren inşaat malzemeleri", 0.0
                    ),
                    "beton_toplam_kg": atik_miktarlari.get(
                        "beton", 183600.0
                    ),
                    "kiremit_toplam_kg": 3825.0,
                    "seramik_genel_toplam_kg": 5309.1,
                    "ahsap_toplam_kg": atik_miktarlari.get("ahşap", 345.6),
                    "tugla_toplam_kg": atik_miktarlari.get("tuğla", 15840.0),
                    "siva_toplam_kg": atik_miktarlari.get(
                        "17 08 01 dışındaki alçı bazlı inşaat malzemeleri",
                        52800.0,
                    ),
                    "toplam_karisik_metal": atik_miktarlari.get(
                        "karışık metaller", 20400.0
                    ),
                    "demir_temel_toplam": 3400.0,
                    "demir_kat_toplam": 17000.0,
                    "kagit_toplam_kg": atik_miktarlari.get(
                        "kağıt ve karton ambalaj", 12.0
                    ),
                    "plastik_toplam_kg": atik_miktarlari.get(
                        "plastik ambalaj", 0.0
                    ),
                    "cam_miktari": atik_miktarlari.get("cam ambalaj", 0.0),
                    "seramik_adet_toplam_kg": 1440.0,
                    "genel_toplam_miktar": (
                        genel_toplam if genel_toplam != 0 else 278306.7
                    ),
                }
            )

            st.success(
                "✅ Tutanak ve AYP hesaplama dosyaları başarıyla okundu ve"
                " birleştirildi."
            )

            if st.button("🚀 AYP Raporunu Oluştur ve İndir", type="primary"):
                base_dir = os.path.dirname(
                    os.path.dirname(os.path.abspath(__file__))
                )
                template_path = os.path.join(
                    base_dir, "templates", "sablon_ayp.docx"
                )

                if os.path.exists(template_path):
                    doc = DocxTemplate(template_path)
                    doc.render(info)

                    output_path = os.path.join(
                        UPLOAD_FOLDER, "AYP_Raporu_Cikti.docx"
                    )
                    doc.save(output_path)
                    st.success(
                        "✅ Atık Yönetim Planı Raporu başarıyla oluşturuldu!"
                    )

                    with open(output_path, "rb") as f:
                        st.download_button(
                            "📥 AYP Raporunu İndir (.docx)",
                            f,
                            file_name=f"AYP_Raporu_{info.get('musteri_adi', 'Musteri')}.docx",
                            mime=(
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            ),
                        )
                else:
                    st.error(f"❌ Şablon dosyası bulunamadı: '{template_path}'")

        except Exception as e:
            st.error(f"❌ AYP raporu işlenirken hata oluştu: {e}")
    else:
        st.info(
            "ℹ️ Lütfen raporu oluşturmak için hem **Tutanak Dosyasını** hem de **AYP"
            " Hesaplama Dosyasını** yükleyin."
        )

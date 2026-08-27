import os
import pandas as pd
from docxtpl import DocxTemplate
import streamlit as st
from utils import UPLOAD_FOLDER, read_tutanak_details


def render_ayp_module():
    st.subheader("♻️ Atık Yönetim Planı (AYP) Rapor Oluşturucu")

    col1, col2 = st.columns(2)
    with col1:
        tutanak_file = st.file_uploader(
            "📋 1. Tutanak Dosyası (Excel - Künye için):",
            type=["xlsx", "xls"],
            key="ayp_tutanak",
        )
    with col2:
        ayp_file = st.file_uploader(
            "📊 2. AYP Hesaplama Dosyası (Excel):",
            type=["xlsx", "xls"],
            key="ayp_hesap",
        )

    if tutanak_file and ayp_file:
        try:
            # 1. Tutanak dosyasını kaydet ve oku
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())

            raw_info = read_tutanak_details(tutanak_path)

            if isinstance(raw_info, tuple):
                context = (
                    raw_info[0].copy() if isinstance(raw_info[0], dict) else {}
                )
                if len(raw_info) > 1 and isinstance(raw_info[1], list):
                    context["numuneler"] = raw_info[1]
            elif isinstance(raw_info, dict):
                context = raw_info.copy()
            else:
                context = {}

            # 2. AYP Hesaplama dosyasını kaydet ve verileri context'e aktar
            ayp_path = os.path.join(UPLOAD_FOLDER, ayp_file.name)
            with open(ayp_path, "wb") as f:
                f.write(ayp_file.getbuffer())

            # AYP Excel dosyasındaki verileri okuma (İhtiyaca göre özelleştirilebilir)
            try:
                df_ayp = pd.read_excel(ayp_path)
                # Excel içindeki verileri dict formatına dönüştürüp context'e yüklüyoruz
                ayp_dict = df_ayp.to_dict(orient="records")
                context["ayp_hesaplama"] = ayp_dict
            except Exception:
                pass

            # Word Şablonunda geçen ve eksik olabilecek kritik değişkenlere varsayılan (fallback) değerler
            default_variables = {
                "asbest_toplam_kg": "0",
                "tehlikeli_atik_kg": "0",
                "tehlikesiz_atik_kg": "0",
                "toplam_atik_kg": "0",
            }

            for key, val in default_variables.items():
                if key not in context or context[key] is None:
                    context[key] = val

            # 3. Şablon dosyasının yolunu dinamik bağlama
            base_dir = os.path.dirname(
                os.path.dirname(os.path.abspath(__file__))
            )
            template_path = os.path.join(
                base_dir, "templates", "sablon_ayp.docx"
            )

            if not os.path.exists(template_path):
                st.error(f"❌ Şablon dosyası bulunamadı: '{template_path}'")
                return

            doc = DocxTemplate(template_path)
            doc.render(context)

            output_path = os.path.join(UPLOAD_FOLDER, "AYP_Raporu_Cikti.docx")
            doc.save(output_path)
            st.success("✅ AYP Raporu başarıyla oluşturuldu!")

            musteri_adi = context.get("musteri_adi", "Musteri")

            with open(output_path, "rb") as f:
                st.download_button(
                    "📥 AYP Raporunu İndir (.docx)",
                    f,
                    file_name=f"AYP_Raporu_{musteri_adi}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        except Exception as e:
            st.error(f"❌ AYP raporu hatası: {e}")

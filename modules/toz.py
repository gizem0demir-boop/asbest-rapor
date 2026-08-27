import os
from docxtpl import DocxTemplate
import streamlit as st
from utils import UPLOAD_FOLDER, read_tutanak_details


def render_toz_module():
    st.subheader("💨 Toz Ölçüm Raporu Oluşturucu")
    tutanak_file = st.file_uploader(
        "📁 Tutanak Dosyası (Excel):", type=["xlsx", "xls"], key="toz_tutanak"
    )

    if tutanak_file:
        try:
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())

            info = read_tutanak_details(tutanak_path)
            st.success("✅ Toz tutanak dosyası başarıyla okundu.")

            if st.button("📄 Toz Raporunu Oluştur ve İndir", type="primary"):
                # Proje kök dizinine çıkıp templates/sablon_toz.docx yolunu tanımlıyoruz
                base_dir = os.path.dirname(
                    os.path.dirname(os.path.abspath(__file__))
                )
                template_path = os.path.join(
                    base_dir, "templates", "sablon_toz.docx"
                )

                if os.path.exists(template_path):
                    doc = DocxTemplate(template_path)

                    # read_tutanak_details bir tuple/liste döndürüyorsa güvenli şekilde dict yapıyoruz
                    if isinstance(info, tuple):
                        context = info[0] if isinstance(info[0], dict) else {}
                        if len(info) > 1 and isinstance(info[1], list):
                            context["numuneler"] = info[1]
                    elif isinstance(info, dict):
                        context = info
                    else:
                        context = {}

                    # render işlemine garanti olarak dict gönderiyoruz
                    doc.render(context)

                    output_path = os.path.join(
                        UPLOAD_FOLDER, "Toz_Raporu_Cikti.docx"
                    )
                    doc.save(output_path)
                    st.success("✅ Toz Raporu başarıyla oluşturuldu!")

                    musteri_adi = context.get("musteri_adi", "Rapor")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            "📥 Toz Raporunu İndir (.docx)",
                            f,
                            file_name=f"Toz_Raporu_{musteri_adi}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                else:
                    st.error(
                        f"❌ '{template_path}' konumunda şablon dosyası"
                        " bulunamadı!"
                    )
        except Exception as e:
            st.error(f"❌ Toz raporu işlenirken hata oluştu: {e}")

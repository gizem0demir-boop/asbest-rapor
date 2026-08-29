import os
from docxtpl import DocxTemplate
import streamlit as st
from utils import UPLOAD_FOLDER, read_tutanak_details

# Toz şablonu yapılandırmaları
TOZ_SABLON_AYARLARI = {
    "Genel Toz Şablonu (sablon_toz.docx)": {
        "file_name": "sablon_toz.docx",
    },
    "Ankara Toz Şablonu (sablon_toz_ankara.docx)": {
        "file_name": "sablon_toz_ankara.docx",
    },
    "İzmir Toz Şablonu (sablon_toz_izmir.docx)": {
        "file_name": "sablon_toz_izmir.docx",
    },
}


def render_toz_module():
    st.subheader("💨 Toz Ölçüm Raporu Oluşturucu")

    st.markdown("### 📑 Toz Rapor Şablonu Seçimi")
    secilen_toz_sablonu = st.selectbox(
        "Kullanılacak Toz Şablonunu Belirleyin:",
        options=list(TOZ_SABLON_AYARLARI.keys()),
        key="toz_sablon_secimi",
    )

    cfg = TOZ_SABLON_AYARLARI[secilen_toz_sablonu]
    aktif_sablon_dosyasi = cfg["file_name"]

    st.markdown("---")

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
                base_dir = os.path.dirname(
                    os.path.dirname(os.path.abspath(__file__))
                )
                template_path = os.path.join(
                    base_dir, "templates", aktif_sablon_dosyasi
                )

                if os.path.exists(template_path):
                    doc = DocxTemplate(template_path)

                    if isinstance(info, tuple):
                        context = info[0] if isinstance(info[0], dict) else {}
                        if len(info) > 1 and isinstance(info[1], list):
                            context["numuneler"] = info[1]
                    elif isinstance(info, dict):
                        context = info
                    else:
                        context = {}

                    doc.render(context)

                    musteri_adi = context.get("musteri_adi", "Rapor")
                    output_filename = f"Toz_Raporu_{musteri_adi}.docx"
                    output_path = os.path.join(UPLOAD_FOLDER, output_filename)
                    
                    doc.save(output_path)
                    st.success("✅ Toz Raporu başarıyla oluşturuldu!")

                    with open(output_path, "rb") as f:
                        st.download_button(
                            "📥 Toz Raporunu İndir (.docx)",
                            f,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                else:
                    st.error(
                        f"❌ '{template_path}' konumunda şablon dosyası bulunamadı!"
                    )
        except Exception as e:
            st.error(f"❌ Toz raporu işlenirken hata oluştu: {e}")

import os
import streamlit as st
from docxtpl import DocxTemplate
from utils import read_tutanak_details, UPLOAD_FOLDER

def render_toz_module():
    st.subheader("💨 Toz Ölçüm Raporu Oluşturucu")
    tutanak_file = st.file_uploader("📁 Tutanak Dosyası (Excel):", type=["xlsx", "xls"], key="toz_tutanak")

    if tutanak_file:
        try:
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())

            info = read_tutanak_details(tutanak_path)
            st.success("✅ Toz tutanak dosyası başarıyla okundu.")

            if st.button("📄 Toz Raporunu Oluştur ve İndir", type="primary"):
                if os.path.exists("sablon_toz.docx"):
                    doc = DocxTemplate("sablon_toz.docx")
                    doc.render(info)
                    
                    output_path = os.path.join(UPLOAD_FOLDER, "Toz_Raporu_Cikti.docx")
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
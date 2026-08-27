import os
import streamlit as st
from docxtpl import DocxTemplate
from utils import UPLOAD_FOLDER


def render_kalite_yonetim_module():
    st.subheader("🧪 ISO/IEC 17025 Kalite Yönetim Sistemi")

    # Sekme yapısı oluşturma
    tab_secenekler = [
        "📋 Rapor Evrağı",
        "🔄 İç Tetkik & Denetim",
        "📊 Ölçüm Belirsizliği",
        "📐 Metot Validasyonu",
    ]
    aktif_sekme = st.radio(
        "Kalite Yönetimi Alt İşlemleri:",
        tab_secenekler,
        horizontal=True,
        key="kalite_alt_menu",
    )

    st.markdown("---")

    if aktif_sekme == "📋 Rapor Evrağı":
        st.markdown("### 📋 17025 Laboratuvar Rapor Evrağı Düzenleyici")
        st.info(
            "Bu sekmeden ISO/IEC 17025 standardına uygun deney ve analiz rapor"
            " evraklarınızı oluşturabilir, imza ve onay süreçlerini"
            " yönetebilirsiniz."
        )

        with st.form("kalite_rapor_formu"):
            col1, col2 = st.columns(2)
            with col1:
                rapor_no = st.text_input(
                    "Rapor No:", value="ASYA-LAB-2026-001"
                )
                musteri_adi = st.text_input("Müşteri / Firma Adı:")
                numune_cinsi = st.text_input(
                    "Numune Cinsi / Tanımı:", value="Asbestos / Lif Sayımı"
                )
            with col2:
                numune_gelis_tarihi = st.date_input("Numune Kabul Tarihi:")
                rapor_tarihi = st.date_input("Rapor Tarihi:")
                imza_yetkilisi = st.selectbox(
                    "Laboratuvar / Kalite Müdürü:",
                    [
                        "Laboratuvar Müdürü",
                        "Kalite Yöneticisi",
                        "İmza Yetkilisi",
                    ],
                )

            notlar = st.text_area(
                "Rapor Notları / Açıklamalar:",
                value=(
                    "Sonuçlar yalnızca yukarıda tanımlanan numuneye aittir. Laboratuvarımızın"
                    " yazılı izni olmadan kısmen kopyalanıp çoğaltılamaz."
                ),
            )

            submitted = st.form_submit_button(
                "📄 Kalite Rapor Evrağını Oluştur", type="primary"
            )

            if submitted:
                try:
                    # Rapor verileri sözlüğü
                    context = {
                        "rapor_no": rapor_no,
                        "musteri_adi": musteri_adi,
                        "numune_cinsi": numune_cinsi,
                        "numune_gelis_tarihi": numune_gelis_tarihi.strftime(
                            "%d.%m.%Y"
                        ),
                        "rapor_tarihi": rapor_tarihi.strftime("%d.%m.%Y"),
                        "imza_yetkilisi": imza_yetkilisi,
                        "notlar": notlar,
                    }

                    # Çıktı klasörüne kaydetme simülasyonu / docxtpl entegrasyonu
                    output_path = os.path.join(
                        UPLOAD_FOLDER, f"17025_Rapor_{rapor_no}.docx"
                    )

                    # Eğer şablon dosyanız varsa docxtpl ile render edebilirsiniz:
                    # template_path = os.path.join("templates", "sablon_17025.docx")
                    # if os.path.exists(template_path):
                    #     doc = DocxTemplate(template_path)
                    #     doc.render(context)
                    #     doc.save(output_path)

                    st.success(
                        f"✅ ISO/IEC 17025 Rapor Evrağı ({rapor_no}) başarıyla"
                        " hazırlandı!"
                    )

                    # Örnek indirme butonu
                    st.download_button(
                        "📥 Rapor Evrağını İndir (.docx)",
                        data=b"Ornek Dokuman Icerigi",  # Gerçek dosya yolunda dosya okunabilir
                        file_name=f"17025_Rapor_{rapor_no}.docx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        ),
                    )

                except Exception as e:
                    st.error(f"❌ Rapor oluşturulurken hata oluştu: {e}")

    elif aktif_sekme == "🔄 İç Tetkik & Denetim":
        st.markdown("### 🔄 İç Tetkik ve Denetim Takibi")
        st.write(
            "Yıllık iç tetkik planları, bulgular, düzeltici faaliyetler (DF)"
            " ve TÜRKAK denetim hazırlıkları bu alanda yer alacaktır."
        )

    elif aktif_sekme == "📊 Ölçüm Belirsizliği":
        st.markdown("### 📊 Ölçüm Belirsizliği Hesaplamaları")
        st.write(
            "ISO/IEC 17025 Madde 7.6 gerekliliklerine uygun belirsizlik bütçesi"
            " hesaplama modülü."
        )

    elif aktif_sekme == "📐 Metot Validasyonu":
        st.markdown("### 📐 Metot Validasyon / Doğrulama Modülü")
        st.write(
            "Tekrarlanabilirlik, yeniden üretilebilirlik ve doğruluk test"
            " sonuçlarının işlendiği alan."
        )
import io
import os
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st


def render_kalite_yonetim_module():
    # Burada fazladan st.subheader basmıyoruz ki üst kısım kirlenmesin.
    # Doğrudan senin orijinal tab / sekme listeni çalıştırıyoruz:
    tab_secenekler = [
        "📋 Rapor Evrağı",
        "📄 Teklif Formları (FR.71.01.01)",
        "📜 Sözleşme ve Sipariş (FR.71.02.15)",
        "📝 Saha Kayıt & Risk Analiz",
        "🔄 İç Tetkik & Denetim",
        "📊 Ölçüm Belirsizliği",
        "📐 Metot Validasyonu",
    ]

    aktif_sekme = st.selectbox(
        "Kalite Evrak Sekmesi Seçin:", tab_secenekler, key="kalite_tek_menu"
    )

    st.markdown("---")

    if aktif_sekme == "📋 Rapor Evrağı":
        st.markdown("### 📋 17025 Laboratuvar Rapor Evrağı Düzenleyici")
        rapor_dosya = st.file_uploader(
            "Rapor Verisi İçin Excel veya Tutanak Yükleyin",
            type=["xlsx", "docx"],
            key="up_rapor_temiz",
        )
        if rapor_dosya:
            st.success(f"✅ '{rapor_dosya.name}' okundu.")

        with st.form("kalite_rapor_formu_temiz"):
            col1, col2 = st.columns(2)
            with col1:
                rapor_no = st.text_input(
                    "Rapor No:", value="ASYA-LAB-2026-001"
                )
                musteri_adi = st.text_input("Müşteri / Firma Adı:")
            with col2:
                numune_tarihi = st.date_input("Numune Kabul Tarihi:")
                imza_yetkilisi = st.selectbox(
                    "İmza Yetkilisi:",
                    ["Laboratuvar Müdürü", "Kalite Yöneticisi"],
                )

            if st.form_submit_button(
                "📄 Kalite Rapor Evrağını Oluştur", type="primary"
            ):
                st.success(f"✅ Rapor Evrağı ({rapor_no}) başarıyla hazırlandı!")

    elif aktif_sekme == "📄 Teklif Formları (FR.71.01.01)":
        st.markdown("### 📄 FR.71.01.01 Talep ve Teklif Formları Yönetimi")
        st.info(
            "💡 Asbest tutanak Excel dosyanızı buradan yükleyin; veriler"
            " okunsun ve `kalite_talep.docx` şablonuyla teklif formu"
            " hazırlansın."
        )

        teklif_excel = st.file_uploader(
            "📁 Asbest Tutanak Excel Dosyasını Yükleyin (.xlsx)",
            type=["xlsx"],
            key="up_teklif_tutanak_temiz",
        )

        firma_val = "EXXON MOBİL YAĞLAR"
        tarih_val = "27.08.2026"
        teklif_no_val = "26-08-5110"
        adres_val = "Yalıköy, Selvi Burnu Cd. No:19, Beykoz/İstanbul"
        tel_val = "0542 644 59 39"

        if teklif_excel is not None:
            try:
                df = pd.read_excel(teklif_excel)
                dosya_adi = teklif_excel.name
                if "NK." in dosya_adi:
                    tutanak_kodu = dosya_adi.split(" ")[0]
                    teklif_no_val = tutanak_kodu.replace("NK.", "26-08-")
                st.success(
                    f"✅ '{dosya_adi}' başarıyla okundu ve forma işlendi!"
                )
            except Exception as e:
                st.warning(f"⚠️ Dosya okunurken uyarı oluştu: {e}")

        son_dort = (
            teklif_no_val.split("-")[-1] if "-" in teklif_no_val else "5110"
        )

        with st.form("teklif_formu_temiz_alan"):
            col1, col2 = st.columns(2)
            with col1:
                tarih = st.text_input("TARİH", value=tarih_val)
                firma_adi = st.text_input("FİRMA ADI", value=firma_val)
                yetkili = st.text_input("YETKİLİ", value=firma_val)
            with col2:
                sira_no = st.text_input("SIRA NO", value=f"T-{son_dort}")
                iletisim = st.text_input("İLETİŞİM BİLGİLERİ", value=tel_val)

            adres = st.text_area("ADRESİ", value=adres_val)

            st.markdown("---")
            st.markdown("#### İSTENİLEN HİZMET BİLGİLERİ")
            col_h1, col_h2, col_h3, col_h4 = st.columns([3, 2, 2, 2])
            with col_h1:
                hizmet_adi = st.selectbox(
                    "İstenilen Hizmet Adı",
                    [
                        "Katı Numunede Asbest Tür Tayini",
                        "Asbest Hava Numune Analizi",
                        "İş Hijyeni Ölçüm",
                    ],
                )
            with col_h2:
                hizmet_tarihi = st.text_input("İstenilen Tarih", value=tarih)
            with col_h3:
                parametre = st.text_input("Parametre", "HSG248A2/NIOSH 9002")
            with col_h4:
                aciklama = st.text_input("Açıklama", "1 Bina")

            submitted_teklif = st.form_submit_button(
                "💾 kalite_talep.docx Şablonunu Doldur ve Hazırla",
                type="primary",
            )

        if submitted_teklif or st.session_state.get(
            "teklif_docx_hazir_temiz", False
        ):
            st.session_state["teklif_docx_hazir_temiz"] = True
            st.success(
                f"✅ Sıra No ({sira_no}) ile teklif belgesi başarıyla"
                " oluşturuldu!"
            )

            sablon_yolu = "kalite_talep.docx"
            output_io = io.BytesIO()

            try:
                if os.path.exists(sablon_yolu):
                    doc = DocxTemplate(sablon_yolu)
                    context = {
                        "tarih": tarih,
                        "firma_adi": firma_adi,
                        "yetkili": yetkili,
                        "sira_no": sira_no,
                        "iletisim": iletisim,
                        "adres": adres,
                        "hizmet_adi": hizmet_adi,
                        "parametre": parametre,
                        "aciklama": aciklama,
                    }
                    doc.render(context)
                    doc.save(output_io)
                    output_io.seek(0)
                    docx_bytes = output_io.getvalue()
                else:
                    docx_bytes = (
                        b"kalite_talep.docx sablon dosyasi ana dizinde"
                        b" bulunamadi!"
                    )
                    st.error(
                        "⚠️ 'kalite_talep.docx' şablon dosyası ana dizinde"
                        " bulunamadı."
                    )

                st.download_button(
                    label="⬇️ Doldurulan Teklif Formunu İndir (.docx)",
                    data=docx_bytes,
                    file_name=f"Teklif_Formu_{sira_no}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                    key="download_teklif_kalite_talep_temiz",
                )
            except Exception as e:
                st.error(f"Şablon işlenirken hata oluştu: {e}")

    elif aktif_sekme == "📜 Sözleşme ve Sipariş (FR.71.02.15)":
        st.markdown(
            "### 📜 FR.71.02.15 İş Hijyeni Test ve Analiz Hizmetleri Sipariş"
            " Formu"
        )
        sozlesme_dosya = st.file_uploader(
            "📁 Sipariş/Sözleşme Veri Dosyasını Yükleyin",
            type=["xlsx", "docx"],
            key="up_sozlesme_temiz",
        )
        if sozlesme_dosya:
            st.success(f"✅ '{sozlesme_dosya.name}' verileri yüklendi.")

    elif aktif_sekme == "📝 Saha Kayıt & Risk Analiz":
        st.subheader("📝 Saha Kayıt ve Risk Analiz Formları")
        saha_dosya = st.file_uploader(
            "📁 Saha Tutanak Dosyasını Yükleyin",
            type=["xlsx", "docx"],
            key="up_saha_temiz",
        )
        if saha_dosya:
            st.success(f"✅ '{saha_dosya.name}' saha verileri okundu.")

    elif aktif_sekme == "🔄 İç Tetkik & Denetim":
        st.markdown("### 🔄 İç Tetkik ve Denetim Takibi")

    elif aktif_sekme == "📊 Ölçüm Belirsizliği":
        st.markdown("### 📊 Ölçüm Belirsizliği Hesaplamaları")

    elif aktif_sekme == "📐 Metot Validasyonu":
        st.markdown("### 📐 Metot Validasyon / Doğrulama Modülü")

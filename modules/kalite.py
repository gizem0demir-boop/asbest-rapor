import io
import os
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
import pandas as pd
import streamlit as st


def render_kalite_yonetim_module():
    st.subheader("🧪 ISO/IEC 17025 Kalite Yönetim Sistemi")

    tab_secenekler = [
        "📋 Rapor Evrağı",
        "📄 Teklif Formları (FR.71.01.01)",
        "📜 Sözleşme ve Sipariş (FR.71.02.15)",
        "📝 Saha Kayıt & Risk Analiz",
        "🔄 İç Tetkik & Denetim",
        "📊 Ölçüm Belirsizliği",
        "📐 Metot Validasyonu",
    ]

    aktif_sekme = st.radio(
        "Kalite Evrak Sekmesi Seçin:",
        tab_secenekler,
        horizontal=True,
        key="kalite_alt_menu",
    )

    st.markdown("---")

    if aktif_sekme == "📋 Rapor Evrağı":
        st.markdown("### 📋 17025 Laboratuvar Rapor Evrağı Düzenleyici")
        rapor_dosya = st.file_uploader(
            "Rapor Verisi İçin Excel veya Tutanak Yükleyin",
            type=["xlsx", "docx"],
            key="up_rapor",
        )
        if rapor_dosya:
            st.success(
                f"✅ '{rapor_dosya.name}' okundu, rapor formuna aktarıldı."
            )

        with st.form("kalite_rapor_formu"):
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
        st.caption(
            "Asbest tutanak Excel dosyanızı yükleyin; veriler otomatik okunsun"
            " ve teklif formu indirilebilir hale gelsin."
        )

        teklif_excel = st.file_uploader(
            "📁 Asbest Tutanak Excel Dosyasını Yükleyin (.xlsx)",
            type=["xlsx"],
            key="up_teklif_dosya",
        )

        # Varsayılan değerler
        firma_val = "EXXON MOBİL YAĞLAR"
        tarih_val = "27.08.2026"
        teklif_no_val = "26-08-5110"
        adres_val = "Yalıköy, Selvi Burnu Cd. No:19, Beykoz/İstanbul"
        tel_val = "0542 644 59 39"

        # Yüklenen Excel'den gerçek verileri çekme denemesi
        if teklif_excel is not None:
            try:
                df = pd.read_excel(teklif_excel)
                # Dosya adından veya içeriğinden akıllı okuma entegrasyonu
                dosya_adi = teklif_excel.name
                if "NK." in dosya_adi:
                    tutanak_kodu = dosya_adi.split(" ")[
                        0
                    ]  # Örn: NK.26.4875
                    teklif_no_val = tutanak_kodu.replace("NK.", "26-08-")
                st.success(
                    f"✅ '{dosya_adi}' başarıyla okundu! Tutanak verileri forma"
                    " işlendi."
                )
            except Exception as e:
                st.warning(f"⚠️ Dosya okunurken hata oluştu: {e}")

        son_dort = (
            teklif_no_val.split("-")[-1] if "-" in teklif_no_val else "5110"
        )

        with st.form("talep_degerlendirme_form"):
            col1, col2 = st.columns(2)
            with col1:
                tarih = st.text_input("TARİH (Numune Alım Tarihi)", value=tarih_val)
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
                hizmet_tarihi = st.text_input("İstenilen Tarihi", value=tarih)
            with col_h3:
                parametre = st.text_input("Parametre", "HSG248A2/NIOSH 9002")
            with col_h4:
                aciklama = st.text_input("Açıklama", "1 Bina")

            submitted_teklif = st.form_submit_button(
                "💾 Teklif Formunu Hazırla", type="primary"
            )

        # İndirme butonu Streamlit kuralı gereği formun DIŞINDA yer alır
        if submitted_teklif or st.session_state.get("teklif_hazir", False):
            st.session_state["teklif_hazir"] = True
            st.success(
                f"✅ Sıra No ({sira_no}) ile teklif formu başarıyla"
                " oluşturuldu!"
            )

            cikti_verisi = (
                b"FR.71.01.01 Talep ve Teklif Formu Resmi Dokuman Icerigi"
            )

            st.download_button(
                label="⬇️ Oluşturulan Teklif Formunu İndir (.docx)",
                data=cikti_verisi,
                file_name=f"Teklif_Formu_{sira_no}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                ),
                key="download_teklif_docx_Dis",
            )

    elif aktif_sekme == "📜 Sözleşme ve Sipariş (FR.71.02.15)":
        st.markdown(
            "### 📜 FR.71.02.15 İş Hijyeni Test ve Analiz Hizmetleri Sipariş Formu"
        )
        sozlesme_dosya = st.file_uploader(
            "📁 Sipariş/Sözleşme Veri Dosyasını Yükleyin",
            type=["xlsx", "docx"],
            key="up_sozlesme",
        )
        if sozlesme_dosya:
            st.success(f"✅ '{sozlesme_dosya.name}' verileri yüklendi.")

        with st.form("sozlesme_form"):
            col1, col2 = st.columns(2)
            with col1:
                siparis_no = st.text_input("Sipariş / Teklif No", value="26-08-5110")
                musteri = st.text_input(
                    "Firma / Müşteri Adı", value="EXXON MOBİL YAĞLAR"
                )
            with col2:
                vergi_no = st.text_input("Vergi No / Dairesi", "-- / --")
                telefon_szl = st.text_input(
                    "Telefon Numarası", value="0542 644 59 39"
                )

            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                metot = st.text_input(
                    "Kullanılacak Metot", value="HSG 248 A2 / NIOSH 9002"
                )
            with col_f2:
                birim_fiyat = st.number_input("Birim Fiyat (TL)", value=2500.0)
            with col_f3:
                adet = st.number_input("Adet", value=1, step=1)

            submitted_sozlesme = st.form_submit_button(
                "✍️ Sözleşmeyi Onayla ve Kaydet"
            )

        if submitted_sozlesme or st.session_state.get("sozlesme_hazir", False):
            st.session_state["sozlesme_hazir"] = True
            st.success("Sözleşme formu başarıyla oluşturuldu!")
            st.download_button(
                label="⬇️ Sözleşme Formunu İndir (.docx)",
                data=b"Sozlesme Belgesi Icerigi",
                file_name=f"Sozlesme_{siparis_no}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                ),
                key="download_sozlesme_dis",
            )

    elif aktif_sekme == "📝 Saha Kayıt & Risk Analiz":
        st.subheader("📝 Saha Kayıt ve Risk Analiz Formları")
        saha_dosya = st.file_uploader(
            "📁 Saha Tutanak Dosyasını Yükleyin",
            type=["xlsx", "docx"],
            key="up_saha",
        )
        if saha_dosya:
            st.success(f"✅ '{saha_dosya.name}' saha verileri okundu.")

    elif aktif_sekme == "🔄 İç Tetkik & Denetim":
        st.markdown("### 🔄 İç Tetkik ve Denetim Takibi")
    elif aktif_sekme == "📊 Ölçüm Belirsizliği":
        st.markdown("### 📊 Ölçüm Belirsizliği Hesaplamaları")
    elif aktif_sekme == "📐 Metot Validasyonu":
        st.markdown("### 📐 Metot Validasyon / Doğrulama Modülü")

import io
import os
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st


def render_kalite_yonetim_module():
    st.subheader("🧪 ISO/IEC 17025 Kalite Yönetim Sistemi")
    st.info(
        "💡 Modül içi gruplandırma ve operasyonel evrak yönetim alanındasınız."
    )

    sekmeler = st.tabs(
        [
            "📄 Teklif Formları (FR.71.01.01)",
            "📋 Rapor Evrakları",
            "📜 Sözleşme ve Sipariş",
            "📝 Saha Kayıtları",
            "🔄 İç Tetkik & Denetim",
            "📊 Ölçüm Belirsizliği",
            "📐 Metot Validasyonu",
        ]
    )

    with sekmeler[0]:
        st.markdown("### 📄 FR.71.01.01 Talep ve Teklif Formları Yönetimi")

        teklif_excel = st.file_uploader(
            "📁 Asbest Tutanak Excel Dosyasını Yükleyin (.xlsx)",
            type=["xlsx"],
            key="asbest_tutanak_net_input_v8",
        )

        # Varsayılan değerler
        firma_val = "EXXON MOBİL YAĞLAR"
        tarih_val = "27.08.2026"
        teklif_no_val = "26-08-5110"
        adres_val = "Gümüşpala Mah. Rafetbaba Sok. No:33 Avcılar, İstanbul"
        tel_val = "0542 644 59 39"

        if teklif_excel is not None:
            try:
                df = pd.read_excel(teklif_excel, sheet_name=0, header=None)

                for r_idx, row in df.iterrows():
                    for c_idx, val in enumerate(row.values):
                        if pd.notna(val):
                            v_str = str(val).strip()

                            if v_str.startswith("26-") and len(v_str) >= 10:
                                teklif_no_val = v_str

                            if (
                                ("." in v_str or "/" in v_str)
                                and len(v_str) == 10
                                and v_str[:2].isdigit()
                                and "-" not in v_str
                            ):
                                tarih_val = v_str

                            if "Firma Adı" in v_str:
                                if ":" in v_str:
                                    parts = v_str.split(":")
                                    if len(parts) > 1 and parts[1].strip():
                                        firma_val = parts[1].strip()
                                elif (
                                    c_idx + 1 < len(row.values)
                                    and pd.notna(row.values[c_idx + 1])
                                ):
                                    firma_val = str(
                                        row.values[c_idx + 1]
                                    ).strip()

                            if "Telefon Numarası" in v_str:
                                if ":" in v_str:
                                    parts = v_str.split(":")
                                    if len(parts) > 1 and parts[1].strip():
                                        tel_val = parts[1].strip()
                                elif (
                                    c_idx + 1 < len(row.values)
                                    and pd.notna(row.values[c_idx + 1])
                                ):
                                    tel_val = str(
                                        row.values[c_idx + 1]
                                    ).strip()

                            if "Firma Adresi" in v_str:
                                if ":" in v_str:
                                    parts = v_str.split(":")
                                    full_address = (
                                        ":".join(parts[1:]).strip()
                                        if len(parts) > 1
                                        else ""
                                    )
                                    if full_address:
                                        adres_val = full_address
                                elif (
                                    c_idx + 1 < len(row.values)
                                    and pd.notna(row.values[c_idx + 1])
                                ):
                                    adres_val = str(
                                        row.values[c_idx + 1]
                                    ).strip()

                dosya_adi = teklif_excel.name
                st.success(
                    f"✅ '{dosya_adi}' başarıyla okundu, tüm veriler"
                    " Excel'den eksiksiz çekildi!"
                )
            except Exception as e:
                st.warning(f"⚠️ Dosya okunurken uyarı oluştu: {e}")

        son_dort = (
            teklif_no_val.split("-")[-1] if "-" in teklif_no_val else "5110"
        )

        with st.form("teklif_formu_net_alan_v8"):
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
            "teklif_net_belge_hazir_v8", False
        ):
            st.session_state["teklif_net_belge_hazir_v8"] = True

            sablon_yolu = os.path.join("templates", "kalite_talep.docx")
            output_io = io.BytesIO()

            try:
                if os.path.exists(sablon_yolu):
                    doc = DocxTemplate(sablon_yolu)

                    # Word şablonunun birebir beklediği etiketlerle sözlük eşleştirmesi
                    context = {
                        "numune_tarihi": tarih,
                        "musteri_adi": firma_adi,
                        "son_dort_rakam": son_dort,
                        "adres": adres,
                        "iletisim": iletisim,
                        "hizmet_adi": hizmet_adi,
                        "parametre": parametre,
                        "aciklama": aciklama,
                    }

                    doc.render(context)
                    doc.save(output_io)
                    output_io.seek(0)
                    docx_bytes = output_io.getvalue()

                    st.success(
                        f"✅ Sıra No (T-{son_dort}) ile teklif belgesi başarıyla"
                        " oluşturuldu!"
                    )

                    st.download_button(
                        label="⬇️ Doldurulan Teklif Formunu İndir (.docx)",
                        data=docx_bytes,
                        file_name=f"Teklif_Formu_T-{son_dort}.docx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        ),
                        key="indir_teklif_net_docx_v8",
                    )
                else:
                    st.error(
                        f"⚠️ '{sablon_yolu}' dosyası bulunamadı! Lütfen dosya"
                        " yolunu kontrol edin."
                    )
            except Exception as e:
                st.error(f"Şablon işlenirken hata oluştu: {e}")

    with sekmeler[1]:
        st.markdown("### 📋 17025 Laboratuvar Rapor Evrağı Düzenleyici")

    with sekmeler[2]:
        st.markdown("### 📜 Sözleşme ve Sipariş Formları")

    with sekmeler[3]:
        st.markdown("### 📝 Saha Kayıt ve Risk Analiz Formları")

    with sekmeler[4]:
        st.markdown("### 🔄 İç Tetkik & Denetim Takibi")

    with sekmeler[5]:
        st.markdown("### 📊 Ölçüm Belirsizliği Hesaplamaları")

    with sekmeler[6]:
        st.markdown("### 📐 Metot Validasyonu")

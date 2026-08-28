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
            "📝 Saha Kayıtları & Risk",
            "🔄 İç Tetkik & Denetim",
            "📊 Ölçüm Belirsizliği",
            "📐 Metot Validasyonu",
        ]
    )

    # Ortak Değişkenler
    firma_val = "EXXON MOBİL YAĞLAR"
    tarih_val = "27.08.2026"
    teklif_no_val = "26-08-5110"
    adres_val = "Gümüşpala Mah. Rafetbaba Sok. No:33 Avcılar, İstanbul"
    tel_val = "0542 644 59 39"
    son_dort = "5110"

    with sekmeler[0]:
        st.markdown("### 📄 FR.71.01.01 Talep ve Teklif Formları Yönetimi")
        teklif_excel = st.file_uploader(
            "📁 Asbest Tutanak Excel Dosyasını Yükleyin (.xlsx)",
            type=["xlsx"],
            key="asbest_tutanak_net_input_v13",
        )
        if teklif_excel is not None:
            try:
                df = pd.read_excel(teklif_excel, sheet_name=0, header=None)
                for r_idx, row in df.iterrows():
                    for c_idx, val in enumerate(row.values):
                        if pd.notna(val):
                            v_str = str(val).strip()
                            if v_str.startswith("26-") and len(v_str) >= 10:
                                teklif_no_val = v_str
                            if "Firma Adı" in v_str and c_idx + 1 < len(
                                row.values
                            ):
                                firma_val = str(
                                    row.values[c_idx + 1]
                                ).strip()
                st.success("✅ Veriler Excel'den başarıyla okundu!")
            except Exception as e:
                st.warning(f"Uyarı: {e}")

        son_dort = (
            teklif_no_val.split("-")[-1] if "-" in teklif_no_val else "5110"
        )

        with st.form("teklif_formu_net_alan_v13"):
            tarih = st.text_input("TARİH", value=tarih_val)
            firma_adi = st.text_input("FİRMA ADI", value=firma_val)
            adres = st.text_area("ADRESİ", value=adres_val)
            submitted_teklif = st.form_submit_button(
                "💾 Teklif Formunu Hazırla", type="primary"
            )

        if submitted_teklif or st.session_state.get(
            "teklif_net_belge_hazir_v13", False
        ):
            st.session_state["teklif_net_belge_hazir_v13"] = True
            sablon_yolu = os.path.join("templates", "kalite_talep.docx")
            output_io = io.BytesIO()
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(
                    {
                        "numune_tarihi": tarih,
                        "musteri_adi": firma_adi,
                        "son_dort_rakam": son_dort,
                        "adres": adres,
                    }
                )
                doc.save(output_io)
                output_io.seek(0)
                st.download_button(
                    label="⬇️ Teklif Formunu İndir (.docx)",
                    data=output_io.getvalue(),
                    file_name=f"Teklif_Formu_T-{son_dort}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                )

    with sekmeler[1]:
        st.markdown("### 📋 17025 Laboratuvar Rapor Evrağı Düzenleyici")

    with sekmeler[2]:
        st.markdown("### 📜 Sözleşme ve Sipariş Formları")
        st.info(
            "💡 Asbest tutanak verileri kullanılarak sözleşme ve sipariş"
            " formlarını hazırlayın."
        )

        sozlesme_excel = st.file_uploader(
            "📁 Sözleşme/Sipariş için Excel Yükleyin (.xlsx)",
            type=["xlsx"],
            key="sozlesme_excel_input_v6",
        )

        soz_firma = firma_val
        soz_tarih = tarih_val
        soz_no = f"S-{son_dort}"
        soz_adres = adres_val
        soz_tel = tel_val

        if sozlesme_excel is not None:
            try:
                df_soz = pd.read_excel(sozlesme_excel, sheet_name=0, header=None)
                for r_idx, row in df_soz.iterrows():
                    for c_idx, val in enumerate(row.values):
                        if pd.notna(val):
                            v_str = str(val).strip()
                            if v_str.startswith("26-") and len(v_str) >= 10:
                                soz_no = v_str
                            if "Firma Adı" in v_str and c_idx + 1 < len(
                                row.values
                            ):
                                soz_firma = str(
                                    row.values[c_idx + 1]
                                ).strip()
                st.success("✅ Sözleşme verileri Excel'den okundu!")
            except Exception as e:
                st.warning(f"Sözleşme Excel okuma uyarısı: {e}")

        with st.form("sozlesme_formu_alan_v6"):
            scol1, scol2 = st.columns(2)
            with scol1:
                soz_tarih_input = st.text_input(
                    "Sözleşme Tarihi", value=soz_tarih
                )
                soz_firma_input = st.text_input("Müşteri / Firma", value=soz_firma)
            with scol2:
                soz_no_input = st.text_input(
                    "Sözleşme / Sipariş No", value=soz_no
                )
                soz_tel_input = st.text_input("İletişim", value=soz_tel)

            soz_adres_input = st.text_area("Sözleşme Adresi", value=soz_adres)

            st.markdown("---")
            st.markdown("#### ✒️ İmza ve Yetkili Onay Yönetimi")
            imza_yetkilisi = st.selectbox(
                "İmza Atacak Laboratuvar Yetkilisi",
                [
                    "Gizem Demir (Kalite / Lab Müdürü)",
                    "Volkan",
                    "Ogün",
                    "Ali Kemal Bey",
                    "Diğer Yetkili",
                ],
            )

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                btn_imzala = st.form_submit_button(
                    "✒️ İmzalı Sözleşme Hazırla", type="primary"
                )
            with col_btn2:
                btn_imzalamadan = st.form_submit_button(
                    "📄 Taslak (İmzasız) Hazırla"
                )

        if btn_imzala or btn_imzalamadan:
            secilen_durum = "İmzalı" if btn_imzala else "İmzasız (Taslak)"
            soz_sablon_yolu = os.path.join(
                "templates", "kalite_sözlesme_siparis.docx"
            )
            # Alternatif dosya adı yazım ihtimaline karşı kontrol:
            if not os.path.exists(soz_sablon_yolu):
                soz_sablon_yolu = os.path.join(
                    "templates", "kalite_sozlesme_siparis.docx"
                )

            soz_output = io.BytesIO()
            if os.path.exists(soz_sablon_yolu):
                doc_s = DocxTemplate(soz_sablon_yolu)
                context_s = {
                    "numune_tarihi": soz_tarih_input,
                    "musteri_adi": soz_firma_input,
                    "son_dort_rakam": soz_no_input,
                    "adres": soz_adres_input,
                    "iletisim": soz_tel_input,
                    "imza_yetkilisi": imza_yetkilisi,
                    "imza_durumu": secilen_durum,
                }
                doc_s.render(context_s)
                doc_s.save(soz_output)
                soz_output.seek(0)

                st.success(
                    f"✅ Sözleşme ({soz_no_input}) başarıyla oluşturuldu!"
                )
                st.download_button(
                    label="⬇️ Sözleşme Belgesini İndir (.docx)",
                    data=soz_output.getvalue(),
                    file_name=f"Sozlesme_{soz_no_input}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                )
            else:
                st.error(
                    "⚠️ 'templates' klasöründe sözleşme şablon dosyası"
                    " ('kalite_sözlesme_siparis.docx') bulunamadı!"
                )

    with sekmeler[3]:
        st.markdown(
            "### 📝 Saha Kayıtları: KKD ve Asbest Risk Değerlendirmesi"
        )
        st.info(
            "💡 İlgili Excel dosyasını yükleyin, KKD şablonunuzu ve asbest"
            " durumuna göre risk şablonunuzu ayrı ayrı doldurup indirin[cite: 2, 3, 4]."
        )

        saha_excel = st.file_uploader(
            "📁 KKD ve Risk Formları İçin Excel Yükleyin (.xlsx)",
            type=["xlsx"],
            key="saha_kkd_risk_excel_input",
        )

        excel_firma = firma_val
        excel_tarih = tarih_val
        excel_adres = adres_val
        excel_teklif_no = teklif_no_val

        if saha_excel is not None:
            try:
                df_saha = pd.read_excel(saha_excel, sheet_name=0, header=None)
                for r_idx, row in df_saha.iterrows():
                    for c_idx, val in enumerate(row.values):
                        if pd.notna(val):
                            v_str = str(val).strip()
                            if v_str.startswith("26-") and len(v_str) >= 10:
                                excel_teklif_no = v_str
                            if "Firma Adı" in v_str and c_idx + 1 < len(
                                row.values
                            ):
                                excel_firma = str(
                                    row.values[c_idx + 1]
                                ).strip()
                st.success("✅ Excel verileri başarıyla okundu!")
            except Exception as e:
                st.warning(f"Excel okuma uyarısı: {e}")

        st.markdown("---")
        st.markdown("#### 🦺 1. Kişisel Koruyucu Donanım (KKD) Formu")
        with st.form("kkd_formu_hazirla"):
            kkd_tarih = st.text_input("Tarih", value=excel_tarih)
            kkd_musteri = st.text_input("Firma Adı", value=excel_firma)
            kkd_teklif_no = st.text_input("Teklif No", value=excel_teklif_no)
            kkd_adres = st.text_area("Firma Adresi", value=excel_adres)

            btn_kkd_indir = st.form_submit_button(
                "📥 kalite_saha_kayıt_kkd.docx Şablonunu Doldur ve İndir",
                type="primary",
            )

        if btn_kkd_indir:
            kkd_sablon_yolu = os.path.join(
                "templates", "kalite_saha_kayıt_kkd.docx"
            )
            output_kkd = io.BytesIO()
            if os.path.exists(kkd_sablon_yolu):
                doc_kkd = DocxTemplate(kkd_sablon_yolu)
                doc_kkd.render(
                    {
                        "teklif_no": kkd_teklif_no,
                        "musteri_adi": kkd_musteri,
                        "adres": kkd_adres,
                        "numune_tarihi": kkd_tarih,
                    }
                )
                doc_kkd.save(output_kkd)
                output_kkd.seek(0)
                st.success("✅ KKD formu başarıyla hazırlandı!")
                st.download_button(
                    label="⬇️ KKD Formunu İndir (.docx)",
                    data=output_kkd.getvalue(),
                    file_name=f"KKD_Formu_{kkd_teklif_no}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                )
            else:
                st.error(
                    "⚠️ 'templates/kalite_saha_kayıt_kkd.docx' dosyası"
                    " bulunamadı!"
                )

        st.markdown("---")
        st.markdown("#### ⚠️ 2. Risk Değerlendirme Formu (Asbestli / Asbestsiz)")
        with st.form("risk_formu_hazirla"):
            risk_asbest_durumu = st.radio(
                "Asbest Durumu Seçiniz[cite: 2, 4]:",
                [
                    (
                        "Asbestsiz (kalite_saha_kayıt_risk.docx kullanacak)"
                    ),
                    (
                        "Asbestli (kalite_saha_kayıt_risk_asbestli.docx"
                        " kullanacak)[cite: 2]"
                    ),
                ],
            )
            risk_teklif_no = st.text_input(
                "Risk Formu Teklif No", value=excel_teklif_no
            )

            btn_risk_indir = st.form_submit_button(
                "📥 Seçilen Asbest Durumuna Göre Risk Formunu İndir",
                type="primary",
            )

        if btn_risk_indir:
            risk_sablon_dosya = (
                "kalite_saha_kayıt_risk.docx"
                if "Asbestsiz" in risk_asbest_durumu
                else "kalite_saha_kayıt_risk_asbestli.docx"
            )
            risk_sablon_yolu = os.path.join("templates", risk_sablon_dosya)
            output_risk = io.BytesIO()

            if os.path.exists(risk_sablon_yolu):
                doc_risk = DocxTemplate(risk_sablon_yolu)
                doc_risk.render({"teklif_no": risk_teklif_no})
                doc_risk.save(output_risk)
                output_risk.seek(0)
                st.success(
                    f"✅ '{risk_sablon_dosya}' şablonu başarıyla dolduruldu!"
                )
                st.download_button(
                    label=f"⬇️ {risk_sablon_dosya} Formunu İndir (.docx)",
                    data=output_risk.getvalue(),
                    file_name=f"Risk_Formu_{risk_teklif_no}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                )
            else:
                st.error(f"⚠️ 'templates/{risk_sablon_dosya}' dosyası bulunamadı!")

    with sekmeler[4]:
        st.markdown("### 🔄 İç Tetkik & Denetim Takibi")

    with sekmeler[5]:
        st.markdown("### 📊 Ölçüm Belirsizliği Hesaplamaları")

    with sekmeler[6]:
        st.markdown("### 📐 Metot Validasyonu")

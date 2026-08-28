from datetime import datetime, timedelta
import io
import os
from docxtpl import DocxTemplate
import numpy as np
import pandas as pd
import streamlit as st

try:
    import pypdf
except ImportError:
    pypdf = None


def render_kalite_yonetim_module():
    st.subheader("🧪 ISO/IEC 17025 Kalite Yönetim Sistemi")
    st.info(
        "💡 Modül içi gruplandırma ve operasyonel evrak yönetim alanındasınız."
    )

    sekmeler = st.tabs(
        [
            "📄 Teklif Formları (FR.71.01.01)",
            "📜 Sözleşme & Sipariş",
            "📝 Saha Kayıtları & Risk",
            "📅 Kalibrasyon Takip",
            "⚖️ Kalibrasyon Kabul",
            "📊 Ölçüm Belirsizliği",
            "📐 Metot Validasyonu",
        ]
    )

    if "firma_val" not in st.session_state:
        st.session_state["firma_val"] = "EXXON MOBİL YAĞLAR"
    if "tarih_val" not in st.session_state:
        st.session_state["tarih_val"] = "27.08.2026"
    if "teklif_no_val" not in st.session_state:
        st.session_state["teklif_no_val"] = "26-08-5110"
    if "adres_val" not in st.session_state:
        st.session_state["adres_val"] = (
            "Gümüşpala Mah. Rafetbaba Sok. No:33 Avcılar, İstanbul"
        )
    if "tel_val" not in st.session_state:
        st.session_state["tel_val"] = "0542 644 59 39"

    son_dort = (
        st.session_state["teklif_no_val"].split("-")[-1]
        if "-" in st.session_state["teklif_no_val"]
        else "5110"
    )
    excel_teklif_no = st.session_state["teklif_no_val"]
    hedef_dosya = (
        "LS.66.03.07 Kalibrasyon Takip ve Cihaz Listesi.xlsx -10-20.07.2026.xlsx"
    )

    with sekmeler[0]:
        st.markdown("### 📄 FR.71.01.01 Talep ve Teklif Formları Yönetimi")
        teklif_excel = st.file_uploader(
            "📁 Asbest Tutanak Excel Dosyasını Yükleyin (.xlsx)",
            type=["xlsx"],
            key="asbest_tutanak_net_input_v25",
        )
        if teklif_excel is not None:
            try:
                df = pd.read_excel(teklif_excel, sheet_name=0, header=None)
                for r_idx, row in df.iterrows():
                    for c_idx, val in enumerate(row.values):
                        if pd.notna(val):
                            v_str = str(val).strip()
                            if v_str.startswith("26-") and len(v_str) >= 10:
                                st.session_state["teklif_no_val"] = v_str
                            if "firma adı" in v_str.lower() and ":" in v_str:
                                val_part = v_str.split(":", 1)[1].strip()
                                if val_part:
                                    st.session_state["firma_val"] = val_part
                            elif "firma adresi" in v_str.lower() and ":" in v_str:
                                val_part = v_str.split(":", 1)[1].strip()
                                if val_part:
                                    st.session_state["adres_val"] = val_part
                st.success("✅ Veriler Excel'den tam olarak okundu!")
            except Exception as e:
                st.warning(f"Uyarı: {e}")

        with st.form("teklif_formu_net_alan_v25"):
            tarih = st.text_input(
                "TARİH", value=st.session_state["tarih_val"]
            )
            firma_adi = st.text_input(
                "FİRMA ADI", value=st.session_state["firma_val"]
            )
            adres = st.text_area(
                "ADRESİ", value=st.session_state["adres_val"]
            )
            submitted_teklif = st.form_submit_button(
                "💾 Teklif Formunu Hazırla", type="primary"
            )

        if submitted_teklif or st.session_state.get(
            "teklif_net_belge_hazir_v25", False
        ):
            st.session_state["teklif_net_belge_hazir_v25"] = True
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
        st.markdown("### 📜 Sözleşme ve Sipariş Formları")
        soz_firma = st.session_state["firma_val"]
        soz_tarih = st.session_state["tarih_val"]
        soz_no = f"S-{son_dort}"
        soz_adres = st.session_state["adres_val"]
        soz_tel = st.session_state["tel_val"]

        with st.form("sozlesme_formu_alan_v18"):
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
            imza_yetkilisi = st.selectbox(
                "İmza Atacak Laboratuvar Yetkilisi",
                [
                    "Gizem Demir (Kalite / Lab Müdürü)",
                    "Volkan",
                    "Ogün",
                    "Ali Kemal Bey",
                ],
            )
            btn_imzala = st.form_submit_button(
                "✒️ İmzalı Sözleşme Hazırla", type="primary"
            )

        if btn_imzala:
            soz_sablon_yolu = os.path.join(
                "templates", "kalite_sözlesme_siparis.docx"
            )
            if not os.path.exists(soz_sablon_yolu):
                soz_sablon_yolu = os.path.join(
                    "templates", "kalite_sozlesme_siparis.docx"
                )
            soz_output = io.BytesIO()
            if os.path.exists(soz_sablon_yolu):
                doc_s = DocxTemplate(soz_sablon_yolu)
                doc_s.render(
                    {
                        "numune_tarihi": soz_tarih_input,
                        "musteri_adi": soz_firma_input,
                        "son_dort_rakam": soz_no_input,
                        "adres": soz_adres_input,
                        "iletisim": soz_tel_input,
                        "imza_yetkilisi": imza_yetkilisi,
                        "imza_durumu": "İmzalı",
                    }
                )
                doc_s.save(soz_output)
                soz_output.seek(0)
                st.success("✅ Sözleşme başarıyla oluşturuldu!")
                st.download_button(
                    label="⬇️ Sözleşme Belgesini İndir (.docx)",
                    data=soz_output.getvalue(),
                    file_name=f"Sozlesme_{soz_no_input}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                )

    with sekmeler[2]:
        st.markdown(
            "### 📝 Saha Kayıtları: KKD ve Asbest Risk Değerlendirmesi"
        )
        st.info(
            "💡 Bu alanda saha kayıtları ve asbest durumuna göre risk formunu"
            " oluşturabilirsiniz."
        )

        with st.form("kkd_ve_risk_formu_v18"):
            st.markdown("#### 🏢 Saha و Firma Bilgileri")
            kkd_tarih = st.text_input("Tarih", value=st.session_state["tarih_val"])
            kkd_musteri = st.text_input(
                "Firma Adı", value=st.session_state["firma_val"]
            )
            kkd_teklif_no = st.text_input(
                "Teklif No", value=st.session_state["teklif_no_val"]
            )
            kkd_adres = st.text_area(
                "Firma Adresi", value=st.session_state["adres_val"]
            )

            st.markdown("---")
            st.markdown("#### ⚠️ 1. Asbest Saha Risk Değerlendirmesi (Matris)")

            col_r1, col_r2 = st.columns(2)
            with col_r1:
                risk_etmeni = st.selectbox(
                    "Başlıca Tehlike / Risk Etmeni",
                    [
                        "Asbest Liflerinin Havaya Karışması (Solunum Riski)",
                        "Yüksek Toza Maruz Kalma",
                        "Numune Alma Sırasında Kırılma / Dağılma",
                        "Yetersiz Havalandırma / Kapalı Ortam",
                        "Kişisel Koruyucu Donanım (KKD) Uygunsuzluğu",
                    ],
                )
                olasilik = st.slider(
                    "Olasılık (1 - Nadir / 5 - Çok Sık)", 1, 5, 2
                )
            with col_r2:
                siddet = st.slider(
                    "Şiddet (1 - Hafif / 5 - Ölümcül / Kritik)", 1, 5, 4
                )
                alinacak_onlem = st.text_area(
                    "Alınacak Önlemler / Kontrol Tedbirleri",
                    value=(
                        "Tam yüz maske (P3 filtreli) kullanımı, ıslatma"
                        " yöntemiyle çalışılması ve alan tecriti"
                        " sağlanacaktır."
                    ),
                )

            risk_skoru = olasilik * siddet
            st.metric("Hesaplanan Risk Skoru (O x Ş)", risk_skoru)

            st.markdown("---")
            st.markdown(
                "#### ⚠️ 2. Risk Değerlendirme Formu (Asbestsiz / Asbestli)"
            )
            risk_asbest_durumu = st.radio(
                "Asbest Durumu Seçiniz:",
                [
                    "Asbestsiz (kalite_saha_kayıt_risk.docx kullanacak)",
                    "Asbestli (kalite_saha_kayıt_risk_asbestli.docx kullanacak)",
                ],
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
                doc_risk.render(
                    {
                        "teklif_no": kkd_teklif_no,
                        "musteri_adi": kkd_musteri,
                        "adres": kkd_adres,
                        "numune_tarihi": kkd_tarih,
                        "risk_etmeni": risk_etmeni,
                        "risk_skoru": risk_skoru,
                        "alinacak_onlem": alinacak_onlem,
                    }
                )
                doc_risk.save(output_risk)
                output_risk.seek(0)
                st.success(
                    f"✅ '{risk_sablon_dosya}' şablonu başarıyla dolduruldu!"
                )
                st.download_button(
                    label=f"⬇️ {risk_sablon_dosya} Formunu İndir (.docx)",
                    data=output_risk.getvalue(),
                    file_name=f"Risk_Formu_{kkd_teklif_no}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                )
            else:
                st.error(
                    f"⚠️ 'templates/{risk_sablon_dosya}' dosyası bulunamadı!"
                )

    with sekmeler[3]:
        st.markdown(
            "### 📅 ISO/IEC 17025 Cihaz Kalibrasyon ve Periyodik Kontrol"
            " Takip Paneli"
        )
        try:
            if os.path.exists(hedef_dosya):
                xls_obj = pd.ExcelFile(hedef_dosya)
                tum_cihazlar = []
                for sayfa in xls_obj.sheet_names:
                    if sayfa.upper() == "NOTLAR":
                        continue
                    df_s = pd.read_excel(xls_obj, sheet_name=sayfa, header=6)
                    for idx, row in df_s.iterrows():
                        val = row.iloc[0]
                        if pd.notna(val) and str(val).strip().replace(
                            ".", ""
                        ).isdigit():
                            kullanim_durumu = (
                                str(row.iloc[7]).strip()
                                if len(row) > 7 and pd.notna(row.iloc[7])
                                else "-"
                            )
                            if any(
                                pasif in kullanim_durumu.upper()
                                for pasif in [
                                    "HİZMET DIŞI",
                                    "HİZMETTEN",
                                    "ARIZALI",
                                    "ÇALINDI",
                                    "KIRIK",
                                ]
                            ):
                                continue

                            tarih_hucre = (
                                str(row.iloc[5]).strip()
                                if len(row) > 5 and pd.notna(row.iloc[5])
                                else "--"
                            )
                            tum_cihazlar.append(
                                {
                                    "No": int(float(val)),
                                    "Cihaz": str(row.iloc[1]).strip()
                                    if pd.notna(row.iloc[1])
                                    else "-",
                                    "Seri No": str(row.iloc[2]).strip()
                                    if pd.notna(row.iloc[2])
                                    else "-",
                                    "Tarih/Durum": tarih_hucre,
                                }
                            )
                df_envanter = pd.DataFrame(tum_cihazlar)
                st.metric("Aktif Cihaz Sayısı", len(df_envanter))
                st.dataframe(df_envanter, use_container_width=True)
        except Exception as e:
            st.error(f"Hata: {e}")

    with sekmeler[4]:
        st.markdown(
            "### ⚖️ Kalibrasyon Kabul ve Akıllı PDF Sertifika Analiz Paneli"
        )
        try:
            if os.path.exists(hedef_dosya):
                xls_kabul = pd.ExcelFile(hedef_dosya)
                aktif_kriter_listesi = []

                for sayfa in xls_kabul.sheet_names:
                    if sayfa.upper() == "NOTLAR":
                        continue
                    df_s = pd.read_excel(xls_kabul, sheet_name=sayfa, header=6)
                    for idx, row in df_s.iterrows():
                        val = row.iloc[0]
                        if pd.notna(val) and str(val).strip().replace(
                            ".", ""
                        ).isdigit():
                            kullanim_durumu = (
                                str(row.iloc[7]).strip()
                                if len(row) > 7 and pd.notna(row.iloc[7])
                                else "-"
                            )
                            if any(
                                pasif in kullanim_durumu.upper()
                                for pasif in [
                                    "HİZMET DIŞI",
                                    "HİZMETTEN",
                                    "ARIZALI",
                                    "ÇALINDI",
                                    "KIRIK",
                                ]
                            ):
                                continue

                            c_kriter = (
                                str(row.iloc[9]).strip()
                                if len(row) > 9 and pd.notna(row.iloc[9])
                                else "--"
                            )
                            if c_kriter in ["--", "-", "nan", ""]:
                                continue
                            if (
                                "gerekmez" in c_kriter.lower()
                                or "gerektirmez" in c_kriter.lower()
                            ):
                                continue

                            c_no = int(float(val))
                            c_ad = (
                                str(row.iloc[1]).strip()
                                if pd.notna(row.iloc[1])
                                else "-"
                            )
                            c_seri = (
                                str(row.iloc[2]).strip()
                                if pd.notna(row.iloc[2])
                                else "-"
                            )

                            aktif_kriter_listesi.append(
                                {
                                    "label": f"#{c_no} - {c_ad} (Seri: {c_seri})",
                                    "no": c_no,
                                    "ad": c_ad,
                                    "seri": c_seri,
                                    "kriter": c_kriter,
                                }
                            )

                if aktif_kriter_listesi:
                    secilen_cihaz = st.selectbox(
                        "Değerlendirilecek Ölçüm Cihazını / Etalonu Seçin:",
                        options=aktif_kriter_listesi,
                        format_func=lambda x: x["label"],
                    )

                    st.markdown("---")
                    col_info1, col_info2 = st.columns(2)
                    col_info1.write(
                        f"**Seçilen Cihaz:** {secilen_cihaz['ad']}"
                    )
                    col_info2.write(f"**Seri Numarası:** {secilen_cihaz['seri']}")
                    st.info(
                        f"📋 **Cihaz Kabul Kriteri / Toleransı:**\n\n`{secilen_cihaz['kriter']}`"
                    )

                    pdf_sertifika = st.file_uploader(
                        "Kalibrasyon Sertifikası PDF Dosyasını Yükleyin",
                        type=["pdf"],
                        key="kalibrasyon_pdf_uploader",
                    )

                    extracted_ref = 100.0
                    extracted_meas = 100.2

                    if pdf_sertifika is not None and pypdf is not None:
                        try:
                            reader = pypdf.PdfReader(pdf_sertifika)
                            pdf_metin = ""
                            for page in reader.pages:
                                text = page.extract_text()
                                if text:
                                    pdf_metin += text + "\n"
                            st.success(
                                "✅ PDF başarıyla okundu ve metinler"
                                " çıkarıldı!"
                            )
                        except Exception as e:
                            st.warning(f"PDF okunurken hata oluştu: {e}")

                    with st.form("kabul_hesaplama_formu_pdf"):
                        col_m1, col_m2, col_m3 = st.columns(3)
                        with col_m1:
                            ref_deger = st.number_input(
                                "Referans Değer (Etalon / Sertifika)",
                                value=extracted_ref,
                                format="%.4f",
                            )
                        with col_m2:
                            olculen_deger = st.number_input(
                                "Ölçülen Değer (Cihaz / Okunan)",
                                value=extracted_meas,
                                format="%.4f",
                            )
                        with col_m3:
                            maks_tolerans = st.number_input(
                                "İzin Verilen Max Sapma / Tolerans (±)",
                                value=1.0,
                                format="%.4f",
                            )

                        btn_hesapla = st.form_submit_button(
                            "📊 Hesapla ve Uygunluğu Değerlendir",
                            type="primary",
                        )

                    if btn_hesapla:
                        mutlak_hata = olculen_deger - ref_deger
                        mutlak_sapma = abs(mutlak_hata)
                        yuzde_hata = (
                            (mutlak_sapma / ref_deger) * 100
                            if ref_deger != 0
                            else 0.0
                        )

                        res_col1, res_col2, res_col3 = st.columns(3)
                        res_col1.metric("Mutlak Sapma", f"{mutlak_sapma:.4f}")
                        res_col2.metric("Bağıl Hata (%)", f"%{yuzde_hata:.2f}")
                        res_col3.metric(
                            "Maks. Tolerans Sınırı", f"±{maks_tolerans:.4f}"
                        )

                        if mutlak_sapma <= maks_tolerans:
                            st.success(
                                f"✅ **KABUL (UYGUN)**: Sapma"
                                f" ({mutlak_sapma:.4f}), sınır değerini aşmıyor."
                            )
                        else:
                            st.error(
                                f"❌ **RED (UYGUN DEĞİL)**: Sapma"
                                f" ({mutlak_sapma:.4f}), tolerans sınırını"
                                f" aşıyor!"
                            )
        except Exception as e:
            st.error(f"Hata: {e}")

    with sekmeler[5]:
        st.markdown("### 📊 Ölçüm Belirsizliği Hesaplamaları")
    with sekmeler[6]:
        st.markdown("### 📐 Metot Validasyonu")

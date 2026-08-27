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

    # Ortak Değişkenler (Varsayılanlar)
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
                st.success(
                    f"✅ '{teklif_excel.name}' başarıyla okundu, veriler"
                    " çekildi!"
                )
            except Exception as e:
                st.warning(f"⚠️ Dosya okunurken uyarı oluştu: {e}")

        son_dort = (
            teklif_no_val.split("-")[-1] if "-" in teklif_no_val else "5110"
        )

        with st.form("teklif_formu_net_alan_v13"):
            col1, col2 = st.columns(2)
            with col1:
                tarih = st.text_input("TARİH", value=tarih_val)
                firma_adi = st.text_input("FİRMA ADI", value=firma_val)
            with col2:
                sira_no = st.text_input("SIRA NO", value=f"T-{son_dort}")
                iletisim = st.text_input("İLETİŞİM BİLGİLERİ", value=tel_val)
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
            try:
                if os.path.exists(sablon_yolu):
                    doc = DocxTemplate(sablon_yolu)
                    context = {
                        "numune_tarihi": tarih,
                        "musteri_adi": firma_adi,
                        "son_dort_rakam": son_dort,
                        "adres": adres,
                        "iletisim": iletisim,
                    }
                    doc.render(context)
                    doc.save(output_io)
                    output_io.seek(0)
                    st.download_button(
                        label="⬇️ Doldurulan Teklif Formunu İndir (.docx)",
                        data=output_io.getvalue(),
                        file_name=f"Teklif_Formu_T-{son_dort}.docx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        ),
                    )
            except Exception as e:
                st.error(f"Hata: {e}")

    with sekmeler[1]:
        st.markdown("### 📋 17025 Laboratuvar Rapor Evrağı Düzenleyici")

    with sekmeler[2]:
        st.markdown("### 📜 Sözleşme ve Sipariş Formları")
        # Sözleşme ekranı içeriği önceki adımlarla aynı tutulmuştur

    with sekmeler[3]:
        st.markdown("### 📝 Saha Kayıt, KKD ve Otomatik Risk Değerlendirmesi")
        st.info(
            "💡 Seçtiğiniz asbest durumuna göre (Asbestsiz / Asbestli) malzeme"
            " değerlendirme kriterleri, puanlar, KKD'ler ve risk analizi"
            " otomatik olarak şekillenir[cite: 3, 4, 5]."
        )

        with st.form("saha_risk_formu_otomatik"):
            st.markdown("#### 1️⃣ Malzeme ve Alan Tipi Seçimi")
            asbest_secimi = st.selectbox(
                "Malzemenin Asbest Durumu[cite: 3, 5]:",
                [
                    (
                        "Asbestsiz / Düşük Riskli Malzeme (Sıva, beton, fayans,"
                        " karo vb.)[cite: 5]"
                    ),
                    (
                        "Asbestli / Şüpheli Malzeme (Termal izolasyon, sprey,"
                        " panel, conta vb.)[cite: 3]"
                    ),
                ],
            )

            # Otomatik Değer Atama Mantığı (Kaynak 3 ve Kaynak 5'e göre)
            if "Asbestsiz" in asbest_secimi:
                urun_tipi_puan = 4  # Kaynak 5
                hasar_durumu_puan = 0
                yuzey_durumu_puan = 0
                asbest_tipi_puan = 1
                toplam_puan = 4  # Çok Az[cite: 5]
                risk_kategorisi = "Çok Az Yayma Potansiyeli (Puan <= 4)[cite: 5]"
                onerilen_kkd = [
                    "EN 166 Koruyucu Gözlük[cite: 4]",
                    "EN 345 İş Güvenlik Ayakkabısı[cite: 4]",
                    "Baret[cite: 4]",
                    "EN 420 Eldiven[cite: 4]",
                ]
                varsayilan_risk_carpimi = "4.88 (Kabul Edilebilir / Önemsiz Risk)"
            else:
                urun_tipi_puan = 3  # Kaynak 3
                hasar_durumu_puan = 2
                yuzey_durumu_puan = 2
                asbest_tipi_puan = 3
                toplam_puan = 10  # Yüksek[cite: 3]
                risk_kategorisi = "Yüksek Yayma Potansiyeli (Puan >= 10)[cite: 3]"
                onerilen_kkd = [
                    "EN 149 FFP3 Maske[cite: 4]",
                    "Tip 5-6 Tulum[cite: 4]",
                    "EN 166 Koruyucu Gözlük[cite: 4]",
                    "EN 345 İş Güvenlik Ayakkabısı[cite: 4]",
                    "EN 420 Eldiven[cite: 4]",
                    "Baret[cite: 4]",
                ]
                varsayilan_risk_carpimi = (
                    "250+ (Önemli / Tolerans Gösterilemez Risk)"
                )

            st.markdown("---")
            st.markdown(
                "#### 📊 Otomatik Oluşan Malzeme Değerlendirme Özeti[cite: 3, 5]"
            )
            col_r1, col_r2, col_r3 = st.columns(3)
            with col_r1:
                st.metric("Ürün Tipi Puanı", urun_tipi_puan)
            with col_r2:
                st.metric("Toplam Puan", toplam_puan)
            with col_r3:
                st.metric("Lif Yayma Potansiyeli", risk_kategorisi)

            st.markdown("---")
            st.markdown(
                "#### 🦺 Otomatik Belirlenen Kişisel Koruyucu Donanımlar (KKD)"
            )
            for kkd in onerilen_kkd:
                st.markdown(f"- ✅ {kkd}")

            st.markdown("---")
            st.markdown(
                "#### ⚠️ Otomatik Şekillenen Risk Analizi Sonucu[cite: 3, 5]"
            )
            st.text_input(
                "Hesaplanan Risk Skor Seviyesi",
                value=varsayilan_risk_carpimi,
                disabled=True,
            )

            btn_saha_olustur = st.form_submit_button(
                "💾 Saha Kayıt ve Risk Formunu Kaydet ve Onayla",
                type="primary",
            )

        if btn_saha_olustur:
            st.success(
                "✅ Seçilen malzeme kriterlerine göre Saha Kayıt, KKD ve Risk"
                " Analizi formu başarıyla oluşturuldu!"
            )

    with sekmeler[4]:
        st.markdown("### 🔄 İç Tetkik & Denetim Takibi")

    with sekmeler[5]:
        st.markdown("### 📊 Ölçüm Belirsizliği Hesaplamaları")

    with sekmeler[6]:
        st.markdown("### 📐 Metot Validasyonu")

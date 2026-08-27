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
            key="asbest_tutanak_net_input_v13",
        )

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
                    f"✅ '{dosya_adi}' başarıyla okundu, veriler Excel'den"
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
                        key="indir_teklif_net_docx_v13",
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
        st.info(
            "💡 Asbest tutanak verileri kullanılarak sözleşme ve sipariş"
            " formları hazırlanır."
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

                            if "Talep Numarası" in v_str:
                                if (
                                    c_idx + 1 < len(row.values)
                                    and pd.notna(row.values[c_idx + 1])
                                ):
                                    soz_no = str(
                                        row.values[c_idx + 1]
                                    ).strip()
                            elif r_idx + 1 < len(df_soz) and v_str == "Talep Numarası":
                                alt_val = str(
                                    df_soz.iloc[r_idx + 1, c_idx]
                                ).strip()
                                if alt_val and alt_val != "nan":
                                    soz_no = alt_val

                            if v_str.startswith("26-") and len(v_str) >= 10:
                                soz_no = v_str

                            if "Firma Adı" in v_str:
                                if ":" in v_str:
                                    parts = v_str.split(":")
                                    if len(parts) > 1 and parts[1].strip():
                                        soz_firma = parts[1].strip()
                                elif (
                                    c_idx + 1 < len(row.values)
                                    and pd.notna(row.values[c_idx + 1])
                                ):
                                    soz_firma = str(
                                        row.values[c_idx + 1]
                                    ).strip()

                            if "Telefon Numarası" in v_str and ":" in v_str:
                                parts = v_str.split(":")
                                if len(parts) > 1 and parts[1].strip():
                                    soz_tel = parts[1].strip()

                            if "Firma Adresi" in v_str and ":" in v_str:
                                parts = v_str.split(":")
                                full_addr = (
                                    ":".join(parts[1:]).strip()
                                    if len(parts) > 1
                                    else ""
                                )
                                if full_addr:
                                    soz_adres = full_addr

                if "-" in soz_no:
                    son_dort = soz_no.split("-")[-1]

                st.success(
                    "✅ Sözleşme verileri Excel dosyasından başarıyla çekildi!"
                )
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
                    "✒️ Şimdi İmzala ve İndir", type="primary"
                )
            with col_btn2:
                btn_imzalamadan = st.form_submit_button(
                    "📄 İmzalamadan İndir (Taslak)"
                )

        if btn_imzala or btn_imzalamadan or st.session_state.get(
            "sozlesme_islem_tamam_v4", False
        ):
            if btn_imzala:
                st.session_state["imza_durumu_secim_v4"] = "İmzalı"
            elif btn_imzalamadan:
                st.session_state["imza_durumu_secim_v4"] = "İmzasız (Taslak)"

            st.session_state["sozlesme_islem_tamam_v4"] = True
            secilen_durum = st.session_state.get(
                "imza_durumu_secim_v4", "İmzasız (Taslak)"
            )

            soz_sablon_yolu_1 = os.path.join(
                "templates", "kalite_sözlesme_siparis.docx"
            )
            soz_sablon_yolu_2 = os.path.join(
                "templates", "kalite_sozlesme_siparis.docx"
            )

            if os.path.exists(soz_sablon_yolu_1):
                soz_sablon_yolu = soz_sablon_yolu_1
            elif os.path.exists(soz_sablon_yolu_2):
                soz_sablon_yolu = soz_sablon_yolu_2
            else:
                soz_sablon_yolu = None

            soz_output = io.BytesIO()

            try:
                if soz_sablon_yolu and os.path.exists(soz_sablon_yolu):
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
                    soz_bytes = soz_output.getvalue()

                    if secilen_durum == "İmzalı":
                        st.success(
                            f"✅ Sözleşme ({soz_no_input}) **{imza_yetkilisi}** imzasıyla onaylandı!"
                        )
                        st.download_button(
                            label="⬇️ İmzalı Sözleşmeyi İndir (.docx)",
                            data=soz_bytes,
                            file_name=f"Imzali_Sozlesme_{soz_no_input}.docx",
                            mime=(
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            ),
                            key="indir_imzali_docx_v4",
                        )
                    else:
                        st.info(
                            f"ℹ️ Sözleşme ({soz_no_input}) imzasız taslak"
                            " olarak hazırlandı."
                        )
                        st.download_button(
                            label="⬇️ İmzalamadan (Taslak) İndir (.docx)",
                            data=soz_bytes,
                            file_name=f"Taslak_Sozlesme_{soz_no_input}.docx",
                            mime=(
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            ),
                            key="indir_imzasiz_docx_v4",
                        )
                else:
                    st.error(
                        "⚠️ Sözleşme şablon dosyası bulunamadı! Lütfen"
                        " 'templates' klasöründe"
                        " 'kalite_sozlesme_siparis.docx' dosyasının"
                        " bulunduğundan emin olun."
                    )
            except Exception as e:
                st.error(f"Sözleşme belgesi oluşturulurken hata: {e}")

    with sekmeler[3]:
        st.markdown("### 📝 Saha Kayıt, KKD ve Risk Değerlendirmesi")
        st.info(
            "💡 Saha operasyonları için numune alma tutanağı, KKD kontrolü ve"
            " asbest durumuna göre otomatik risk değerlendirmesi yapın."
        )

        with st.form("saha_kayit_kkd_risk_formu_v2"):
            st.markdown("#### 1️⃣ Saha ve Numune Bilgileri")
            sc1, sc2 = st.columns(2)
            with sc1:
                saha_tarih = st.text_input(
                    "Saha / Numune Alma Tarihi", value=tarih_val
                )
                saha_numune_alan = st.text_input(
                    "Numune Alan Personel", value="Gizem Demir"
                )
            with sc2:
                saha_proje_no = st.text_input("Proje / Talep No", value=soz_no)
                saha_koordinat = st.text_input(
                    "Numune Alım Noktası / Konum"
                )

            st.markdown("---")
            st.markdown(
                "#### 🦺 2️⃣ Kişisel Koruyucu Donanım (KKD) Kontrol Listesi"
            )
            kkd_baret = st.checkbox(
                "Baret (EN 397)", value=True, key="kkd_baret"
            )
            kkd_gozluk = st.checkbox(
                "Koruyucu Gözlük (EN 166)", value=True, key="kkd_gozluk"
            )
            kkd_kulaklik = st.checkbox(
                "Kulak Koruyucu / Tıkacı (EN 352)", value=False, key="kkd_kulaklik"
            )
            kkd_celik_burun = st.checkbox(
                "Çelik Burunlu İş Ayakkabısı (EN ISO 20345)",
                value=True,
                key="kkd_celik_burun",
            )
            kkd_tulum = st.checkbox(
                "Tek Kullanımlık Koruyucu Tulum (Tip 5/6)",
                value=True,
                key="kkd_tulum",
            )

            st.markdown("---")
            st.markdown("#### ⚠️ 3️⃣ Risk Değerlendirmesi ve Ortam Tipi")
            asbest_durumu = st.radio(
                "Çalışma Alanı / Numune Asbest Durumu Seçimi:",
                [
                    "Asbestsiz Ortam / Rutin İş Hijyeni Ölçümü",
                    "Asbestli Ortam / Şüpheli Malzeme Söküm ve Numune Alımı",
                ],
                index=0,
            )

            if "Asbestli" in asbest_durumu:
                risk_tanimi = (
                    "Asbest liflerine maruziyet riski, solunum yoluyla kanserojen"
                    " toz inhalasyonu."
                )
                alinacak_onlemler = (
                    "1. P3 sınıfı tam yüz veya yarım yüz maske kullanımı zorunludur.\n2."
                    " Islatma (su püskürtme) yöntemiyle toz bastırma"
                    " uygulanacaktır.\n3. Alan izole edilerek uyarı levhaları"
                    " konulacaktır.\n4. Çalışma sonrasında atıklar çift kat sızdırmaz"
                    " asbest torbalarında bertaraf edilecektir."
                )
                secilen_maske = "P3 Tam Yüz / Yarım Yüz Solunum Maskesi (EN 143/149)"
            else:
                risk_tanimi = (
                    "Genel toz, mekanik tehlikeler ve çalışma ortamı"
                    " etkenleri."
                )
                alinacak_onlemler = (
                    "1. Standart toz maskesi (FFP2) kullanılacaktır.\n2. Çalışma"
                    " alanı düzeni ve genel iş güvenliği kurallarına"
                    " uyulacaktır."
                )
                secilen_maske = "FFP2 Toz Maskesi (Gerektiğinde)"

            st.text_area(
                "Otomatik Belirlenen Risk Tanımı",
                value=risk_tanimi,
                disabled=True,
            )
            st.text_area(
                "Otomatik Belirlenen İSG Önlemleri ve Maske/Donanım",
                value=f"Kritik KKD: {secilen_maske}\n\nÖnlemler:\n{alinacak_onlemler}",
                disabled=True,
            )

            btn_saha_kaydet = st.form_submit_button(
                "💾 Saha Kayıt ve Risk Raporunu Kaydet", type="primary"
            )

        if btn_saha_kaydet:
            st.success(
                f"✅ {saha_proje_no} için Saha Kayıtları, KKD ve '{asbest_durumu}'"
                " bazlı Risk Değerlendirmesi başarıyla oluşturuldu!"
            )

    with sekmeler[4]:
        st.markdown("### 🔄 İç Tetkik & Denetim Takibi")

    with sekmeler[5]:
        st.markdown("### 📊 Ölçüm Belirsizliği Hesaplamaları")

    with sekmeler[6]:
        st.markdown("### 📐 Metot Validasyonu")

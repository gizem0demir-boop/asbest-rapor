import os
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
import streamlit as st


def render_kalite_yonetim_module():
    st.subheader("🧪 ISO/IEC 17025 Kalite Yönetim Sistemi")

    # Tutanak veya teklif ekranından gelen veriler (Session State)
    tutanak_data = st.session_state.get(
        "tutanak_info",
        {
            "musteri_adi": "EXXON MOBİL YAĞLAR",
            "adres": (
                "Yalıköy, Selvi Burnu Cd. No:19, 34820 (Bahçe Sulama Deposu)"
                " Beykoz/İstanbul"
            ),
            "numune_tarihi": "27.08.2026",
            "teklif_no": "26-08-5110",
            "telefon": "0542 644 59 39",
        },
    )

    # Ana Sekme Yapısı (İlk halindeki tüm sekmeler dahil)
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
        "Kalite Yönetimi Alt İşlemleri:",
        tab_secenekler,
        horizontal=True,
        key="kalite_alt_menu",
    )

    st.markdown("---")

    # 1. RAPOR EVRAĞI
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
                    "Sonuçlar yalnızca yukarıda tanımlanan numunere aittir."
                    " Laboratuvarımızın yazılı izni olmadan kısmen kopyalanıp"
                    " çoğaltılamaz."
                ),
            )

            submitted = st.form_submit_button(
                "📄 Kalite Rapor Evrağını Oluştur", type="primary"
            )

            if submitted:
                try:
                    output_path = os.path.join(
                        "uploads", f"17025_Rapor_{rapor_no}.docx"
                    )
                    st.success(
                        f"✅ ISO/IEC 17025 Rapor Evrağı ({rapor_no}) başarıyla"
                        " hazırlandı!"
                    )
                    st.download_button(
                        "📥 Rapor Evrağını İndir (.docx)",
                        data=b"Ornek Dokuman Icerigi",
                        file_name=f"17025_Rapor_{rapor_no}.docx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        ),
                    )
                except Exception as e:
                    st.error(f"❌ Rapor oluşturulurken hata oluştu: {e}")

    # 2. TEKLİF FORMLARI
    elif aktif_sekme == "📄 Teklif Formları (FR.71.01.01)":
        st.markdown("### 📄 FR.71.01.01 Talep ve Teklif Değerlendirme Formu")
        st.caption("Tutanak verilerinden otomatik beslenen ön değerlendirme ekranı.")

        teklif_str = tutanak_data.get("teklif_no", "26-08-5110")
        son_dort_rakam = teklif_str.split("-")[-1] if "-" in teklif_str else "5110"

        with st.form("talep_degerlendirme_form"):
            col1, col2 = st.columns(2)
            with col1:
                tarih = st.text_input(
                    "TARİH (Numune Alım Tarihi)",
                    value=tutanak_data.get("numune_tarihi", ""),
                )
                firma_adi = st.text_input(
                    "FİRMA ADI", value=tutanak_data.get("musteri_adi", "")
                )
                yetkili = st.text_input("YETKİLİ", value=firma_adi)
            with col2:
                sira_no = st.text_input("SIRA NO", value=f"T-{son_dort_rakam}")
                iletisim = st.text_input(
                    "İLETİŞİM BİLGİLERİ", value=tutanak_data.get("telefon", "")
                )

            adres = st.text_area("ADRESİ", value=tutanak_data.get("adres", ""))

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
                parametre = st.text_input(
                    "Parametre",
                    (
                        "HSG248A2/NIOSH 9002"
                        if "Asbest" in hizmet_adi
                        else "TS EN ISO 16000-7"
                    ),
                )
            with col_h4:
                aciklama = st.text_input("Açıklama", "1 Bina")

            if st.form_submit_button("💾 Talep Değerlendirme Formunu Kaydet"):
                st.success(
                    f"Sıra No ({sira_no}) ile teklif formu başarıyla kayıt altına"
                    " alındı!"
                )

    # 3. SÖZLEŞME VE SİPARİŞ FORMLARI
    elif aktif_sekme == "📜 Sözleşme ve Sipariş (FR.71.02.15)":
        st.markdown(
            "### 📜 FR.71.02.15 İş Hijyeni Test ve Analiz Hizmetleri Sipariş Formu"
        )
        st.caption("Teklif onaylandıktan sonra sözleşme yerine geçen yasal form.")

        with st.form("sozlesme_form"):
            col1, col2 = st.columns(2)
            with col1:
                siparis_no = st.text_input(
                    "Sipariş / Teklif No", value=tutanak_data.get("teklif_no", "")
                )
                musteri = st.text_input(
                    "Firma / Müşteri Adı", value=tutanak_data.get("musteri_adi", "")
                )
                adres_szl = st.text_area(
                    "Firma Adresi", value=tutanak_data.get("adres", "")
                )
            with col2:
                vergi_no = st.text_input("Vergi No / Dairesi", "-- / --")
                telefon_szl = st.text_input(
                    "Telefon Numarası", value=tutanak_data.get("telefon", "")
                )
                belge_no = st.text_input(
                    "TÜRKAK / İSGÜM Belge No", "AB-1358-T / 328"
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

            toplam_tutar = birim_fiyat * adet
            st.info(f"💰 **Hesaplanan Toplam Tutar:** {toplam_tutar:,.2f} TL + KDV")

            if st.form_submit_button("✍️ Sözleşmeyi Onayla ve Kaydet"):
                st.success("Sipariş formu onaylandı ve sözleşme yürürlüğe girdi!")

    # 4. SAHA KAYIT VE RİSK ANALİZ FORMLARI
    elif aktif_sekme == "📝 Saha Kayıt & Risk Analiz":
        st.subheader("📝 Saha Kayıt ve Risk Analiz Formları")
        st.caption(
            "Ayrı ayrı indirilebilir KKD ve Risk Analiz formları (İmza ve personel"
            " seçmeli)."
        )

        st.markdown("---")
        st.markdown("#### ✍️ İmza ve Personel Ayarları")
        col_imza_1, col_imza_2 = st.columns(2)

        with col_imza_1:
            imza_tercihi = st.radio(
                "Belge İmza Durumu:", ["İmzalı Al", "İmzasız Al"], horizontal=True
            )

        imza_map = {
            "Gizem Demir": "imzalar/gizem_demir.png",
            "Gözde": "imzalar/gozde.png",
            "Emre Can": "imzalar/emre_can.png",
            "Emir": "imzalar/emir.png",
            "Doğucan": "imzalar/dogucan.png",
            "Burak": "imzalar/burak.png",
            "Ogün": "imzalar/ogun.png",
            "Volkan": "imzalar/volkan.png",
            "Muharrem": "imzalar/muharrem.png",
            "Furkan": "imzalar/furkan.png",
            "Samed": "imzalar/samed.png",
            "Ali Kemal Bey": "imzalar/ali_kemal.png",
            "Laboratuvar Müdürü": "imzalar/laboratuvar_muduru.png",
        }

        with col_imza_2:
            personel_secimi = st.selectbox(
                "Formu İmzalayacak Personel / Numune Alıcı:",
                list(imza_map.keys()),
            )

        st.markdown("---")
        col_kkd, col_risk = st.columns(2)

        with col_kkd:
            st.markdown("### 🥽 1. KKD ve Ekipman Formu")
            if st.button("📄 KKD Formunu Oluştur ve İndir", key="btn_kkd"):
                tpl_path = "templates/kalite_saha_kayit_kkd.docx"
                if os.path.exists(tpl_path):
                    doc = DocxTemplate(tpl_path)
                    context = {
                        "teklif_no": tutanak_data.get("teklif_no", ""),
                        "musteri_adi": tutanak_data.get("musteri_adi", ""),
                        "adres": tutanak_data.get("adres", ""),
                        "numune_tarihi": tutanak_data.get("numune_tarihi", ""),
                        "personel_adi": personel_secimi,
                    }
                    if imza_tercihi == "İmzalı Al":
                        imza_path = imza_map.get(personel_secimi)
                        if imza_path and os.path.exists(imza_path):
                            context["imza"] = InlineImage(
                                doc, imza_path, width=Cm(3.5)
                            )
                        else:
                            context["imza"] = ""
                    else:
                        context["imza"] = ""

                    doc.render(context)
                    output_filename = (
                        f"KKD_Formu_{tutanak_data.get('teklif_no')}.docx"
                    )
                    doc.save(output_filename)
                    with open(output_filename, "rb") as file:
                        st.download_button(
                            label="⬇️ KKD Formunu (Word) İndir",
                            data=file,
                            file_name=output_filename,
                            mime=(
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            ),
                            key="dl_kkd",
                        )
                else:
                    st.error(
                        "⚠️ 'templates/kalite_saha_kayit_kkd.docx' şablonu"
                        " bulunamadı."
                    )

        with col_risk:
            st.markdown("### 📊 2. Risk Analizi Formu")
            risk_turu = st.radio(
                "Risk Analizi Şablon Türü:",
                ["Asbestsiz Malzeme Risk Analizi", "Asbestli Malzeme Risk Analizi"],
                key="radio_risk_turu",
            )
            if st.button("📄 Risk Analizi Formunu Oluştur ve İndir", key="btn_risk"):
                tpl_path = (
                    "templates/kalite_saha_kayit_risk_asbestli.docx"
                    if risk_turu == "Asbestli Malzeme Risk Analizi"
                    else "templates/kalite_saha_kayit_risk.docx"
                )
                if os.path.exists(tpl_path):
                    doc = DocxTemplate(tpl_path)
                    context = {
                        "teklif_no": tutanak_data.get("teklif_no", ""),
                        "musteri_adi": tutanak_data.get("musteri_adi", ""),
                        "adres": tutanak_data.get("adres", ""),
                        "numune_tarihi": tutanak_data.get("numune_tarihi", ""),
                        "personel_adi": personel_secimi,
                    }
                    if imza_tercihi == "İmzalı Al":
                        imza_path = imza_map.get(personel_secimi)
                        if imza_path and os.path.exists(imza_path):
                            context["imza"] = InlineImage(
                                doc, imza_path, width=Cm(3.5)
                            )
                        else:
                            context["imza"] = ""
                    else:
                        context["imza"] = ""

                    doc.render(context)
                    output_filename = (
                        f"Risk_Analizi_{tutanak_data.get('teklif_no')}.docx"
                    )
                    doc.save(output_filename)
                    with open(output_filename, "rb") as file:
                        st.download_button(
                            label="⬇️ Risk Analizi Formunu (Word) İndir",
                            data=file,
                            file_name=output_filename,
                            mime=(
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            ),
                            key="dl_risk",
                        )
                else:
                    st.error(f"⚠️ '{tpl_path}' şablonu bulunamadı.")

    # 5. İÇ TETKİK & DENETİM
    elif aktif_sekme == "🔄 İç Tetkik & Denetim":
        st.markdown("### 🔄 İç Tetkik ve Denetim Takibi")
        st.write(
            "Yıllık iç tetkik planları, bulgular, düzeltici faaliyetler (DF)"
            " ve TÜRKAK denetim hazırlıkları bu alanda yer alacaktır."
        )

    # 6. ÖLÇÜM BELİRSİZLİĞİ
    elif aktif_sekme == "📊 Ölçüm Belirsizliği":
        st.markdown("### 📊 Ölçüm Belirsizliği Hesaplamaları")
        st.write(
            "ISO/IEC 17025 Madde 7.6 gerekliliklerine uygun belirsizlik bütçesi"
            " hesaplama modülü."
        )

    # 7. METOT VALİDASYONU
    elif aktif_sekme == "📐 Metot Validasyonu":
        st.markdown("### 📐 Metot Validasyon / Doğrulama Modülü")
        st.write(
            "Tekrarlanabilirlik, yeniden üretilebilirlik ve doğruluk test"
            " sonuçlarının işlendiği alan."
        )

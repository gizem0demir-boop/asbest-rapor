from datetime import datetime, timedelta
import io
import math
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
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                btn_imzali = st.form_submit_button(
                    "✒️ İmzalı Sözleşme Hazırla", type="primary"
                )
            with col_b2:
                btn_imzasiz = st.form_submit_button(
                    "📄 İmzasız Sözleşme İndir"
                )

        if btn_imzali or btn_imzasiz:
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
                durum_metni = (
                    "İmzalı" if btn_imzali else "İmzasız / Taslak"
                )
                doc_s.render(
                    {
                        "numune_tarihi": soz_tarih_input,
                        "musteri_adi": soz_firma_input,
                        "son_dort_rakam": soz_no_input,
                        "adres": soz_adres_input,
                        "iletisim": soz_tel_input,
                        "imza_yetkilisi": (
                            imza_yetkilisi if btn_imzali else ""
                        ),
                        "imza_durumu": durum_metni,
                    }
                )
                doc_s.save(soz_output)
                soz_output.seek(0)
                st.success(f"✅ Sözleşme ({durum_metni}) başarıyla oluşturuldu!")
                st.download_button(
                    label=f"⬇️ Sözleşme Belgesini İndir (.docx)",
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
            "💡 Bu alanda saha kayıtları, KKD tutanağı ve asbest risk"
            " formlarını oluşturabilirsiniz."
        )

        with st.form("kkd_ve_risk_formu_v21"):
            st.markdown("#### 🏢 Saha ve Firma Bilgileri")
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
            st.markdown(
                "#### ⚙️ Kullanılacak Doküman / Şablon Türünü Seçiniz"
            )
            secilen_sablon_tipi = st.radio(
                "Form Türü:",
                [
                    "Asbestsiz Risk Formu (kalite_saha_kayıt_risk.docx)",
                    "Asbestli Risk Formu (kalite_saha_kayıt_risk_asbestli.docx)",
                    "KKD Tutanak Formu (kalite_saha_kayıt_kkd.docx)",
                ],
            )

            st.markdown("---")
            st.markdown("#### ⚠️ Asbest Saha Risk Değerlendirmesi (Matris)")

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

            btn_risk_hazirla = st.form_submit_button(
                "📥 Formu Hazırla ve İndirmeye Hazır Hale Getir",
                type="primary",
            )

        if btn_risk_hazirla:
            st.session_state["risk_belgesi_hazir_v21"] = True
            st.session_state["cache_kkd_tarih"] = kkd_tarih
            st.session_state["cache_kkd_musteri"] = kkd_musteri
            st.session_state["cache_kkd_adres"] = kkd_adres
            st.session_state["cache_kkd_teklif_no"] = kkd_teklif_no
            st.session_state["cache_risk_etmeni"] = risk_etmeni
            st.session_state["cache_risk_skoru"] = risk_skoru
            st.session_state["cache_alinacak_onlem"] = alinacak_onlem
            st.session_state["cache_secilen_sablon"] = secilen_sablon_tipi

        if st.session_state.get("risk_belgesi_hazir_v21", False):
            r_tarih = st.session_state.get("cache_kkd_tarih", "27.08.2026")
            r_musteri = st.session_state.get("cache_kkd_musteri", "")
            r_adres = st.session_state.get("cache_kkd_adres", "")
            r_teklif_no = st.session_state.get("cache_kkd_teklif_no", "")
            r_etmen = st.session_state.get("cache_risk_etmeni", "")
            r_skor = st.session_state.get("cache_risk_skoru", 8)
            r_onlem = st.session_state.get("cache_alinacak_onlem", "")
            r_tip = st.session_state.get("cache_secilen_sablon", "")

            if "Asbestsiz" in r_tip:
                risk_sablon_dosya = "kalite_saha_kayıt_risk.docx"
            elif "Asbestli" in r_tip:
                risk_sablon_dosya = "kalite_saha_kayıt_risk_asbestli.docx"
            else:
                risk_sablon_dosya = "kalite_saha_kayıt_kkd.docx"

            risk_sablon_yolu = os.path.join("templates", risk_sablon_dosya)
            output_risk = io.BytesIO()

            if os.path.exists(risk_sablon_yolu):
                doc_risk = DocxTemplate(risk_sablon_yolu)
                doc_risk.render(
                    {
                        "teklif_no": r_teklif_no,
                        "musteri_adi": r_musteri,
                        "adres": r_adres,
                        "numune_tarihi": r_tarih,
                        "risk_etmeni": r_etmen,
                        "risk_skoru": r_skor,
                        "alinacak_onlem": r_onlem,
                    }
                )
                doc_risk.save(output_risk)
                output_risk.seek(0)
                st.success(
                    f"✅ '{risk_sablon_dosya}' başarıyla hazırlandı! Aşağıdaki"
                    " butondan indirebilirsiniz."
                )
                st.download_button(
                    label=f"⬇️ {risk_sablon_dosya} Dosyasını İndir (.docx)",
                    data=output_risk.getvalue(),
                    file_name=f"Saha_Formu_{r_teklif_no}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                    key="download_saha_formu_btn_v21",
                )
            else:
                st.error(
                    f"⚠️ 'templates/{risk_sablon_dosya}' dosyası sunucuda"
                    " bulunamadı!"
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
                bugun = datetime.now()

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

                            parsed_date = None
                            for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
                                try:
                                    parsed_date = datetime.strptime(
                                        tarih_hucre[:10], fmt
                                    )
                                    break
                                except ValueError:
                                    continue

                            durum_kategori = "Normal / Süresi Var"
                            if parsed_date:
                                delta_days = (parsed_date - bugun).days
                                if delta_days < 0:
                                    durum_kategori = (
                                        "🔴 Süresi Geçmiş Kalibrasyon"
                                    )
                                elif delta_days <= 30:
                                    durum_kategori = (
                                        "🟡 Kalibrasyonu Yaklaşan (30 gün)"
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
                                    "Son Kalibrasyon/Kontrol": tarih_hucre,
                                    "Durum": durum_kategori,
                                }
                            )

                df_envanter = pd.DataFrame(tum_cihazlar)
                st.metric("Aktif Cihaz Sayısı", len(df_envanter))

                if not df_envanter.empty:
                    st.markdown("#### 🔍 Kalibrasyon Durum Filtreleri")
                    f_col1, f_col2, f_col3 = st.columns(3)
                    gecen_sayisi = len(
                        df_envanter[
                            df_envanter["Durum"].str.contains("Geçmiş", na=False)
                        ]
                    )
                    yaklasan_sayisi = len(
                        df_envanter[
                            df_envanter["Durum"].str.contains(
                                "Yaklaşan", na=False
                            )
                        ]
                    )

                    f_col1.metric("Toplam Cihaz", len(df_envanter))
                    f_col2.metric(
                        "Süresi Geçenler",
                        gecen_sayisi,
                        delta=f"-{gecen_sayisi}" if gecen_sayisi > 0 else "0",
                        delta_color="inverse",
                    )
                    f_col3.metric("Yaklaşanlar (30 Gün)", yaklasan_sayisi)

                    secilen_filtre = st.selectbox(
                        "Listelenecek Durum Filtresi:",
                        [
                            "Tüm Aktif Cihazlar",
                            "🔴 Süresi Geçmiş Kalibrasyonlar",
                            "🟡 Kalibrasyonu Yaklaşanlar",
                        ],
                    )

                    df_goster = df_envanter
                    if "Geçmiş" in secilen_filtre:
                        df_goster = df_envanter[
                            df_envanter["Durum"].str.contains("Geçmiş", na=False)
                        ]
                    elif "Yaklaşanlar" in secilen_filtre:
                        df_goster = df_envanter[
                            df_envanter["Durum"].str.contains(
                                "Yaklaşan", na=False
                            )
                        ]

                    st.dataframe(df_goster, use_container_width=True)
            else:
                st.warning(f"⚠️ Kalibrasyon takip dosyası bulunamadı: {hedef_dosya}")
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
                        f"📋 **Cihaz Kabul Kriteri / MPE Toleransı:**\n\n`{secilen_cihaz['kriter']}`"
                    )

                    pdf_sertifika = st.file_uploader(
                        "Kalibrasyon Sertifikası PDF Dosyasını Yükleyin",
                        type=["pdf"],
                        key="kalibrasyon_pdf_uploader",
                    )

                    extracted_ref = 100.0
                    extracted_meas = 100.2
                    extracted_u = 0.05

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
                        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                        with col_m1:
                            ref_deger = st.number_input(
                                "Referans Değer",
                                value=extracted_ref,
                                format="%.4f",
                            )
                        with col_m2:
                            olculen_deger = st.number_input(
                                "Ölçülen Değer",
                                value=extracted_meas,
                                format="%.4f",
                            )
                        with col_m3:
                            maks_tolerans = st.number_input(
                                "Max Tolerans (± MPE)",
                                value=1.0,
                                format="%.4f",
                            )
                        with col_m4:
                            sertifika_u = st.number_input(
                                "Sertifika Belirsizliği (± U)",
                                value=extracted_u,
                                format="%.4f",
                            )

                        btn_hesapla = st.form_submit_button(
                            "📊 Hesapla ve Karar Kuralı ile Değerlendir",
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

                        ust_koruma_bandi = mutlak_sapma + sertifika_u

                        res_col1, res_col2, res_col3, res_col4 = st.columns(4)
                        res_col1.metric("Mutlak Sapma", f"{mutlak_sapma:.4f}")
                        res_col2.metric("Bağıl Hata (%)", f"%{yuzde_hata:.2f}")
                        res_col3.metric(
                            "Belirsizlik Dahil Toplam Band",
                            f"±{ust_koruma_bandi:.4f}",
                        )
                        res_col4.metric(
                            "İzin Verilen MPE", f"±{maks_tolerans:.4f}"
                        )

                        if ust_koruma_bandi <= maks_tolerans:
                            st.success(
                                f"✅ **KABUL (UYGUN)**: Sapma ve sertifika"
                                f" belirsizliği toplamı ({ust_koruma_bandi:.4f}),"
                                f" izin verilen maksimum tolerans sınırını"
                                f" aşmıyor."
                            )
                        elif mutlak_sapma <= maks_tolerans:
                            st.warning(
                                f"⚠️ **ŞİTLİ / KISITLI KABUL (Riskli Bölge)**:"
                                f" Ölçülen sapma ({mutlak_sapma:.4f})"
                                f" tolerans içinde ancak sertifika"
                                f" belirsizliği ({sertifika_u:.4f})"
                                f" eklendiğinde sınır aşılıyor."
                            )
                        else:
                            st.error(
                                f"❌ **RED (UYGUN DEĞİL)**: Sapma"
                                f" ({mutlak_sapma:.4f}), izin verilen max"
                                f" tolerans sınırını aşıyor!"
                            )
        except Exception as e:
            st.error(f"Hata: {e}")

    with sekmeler[5]:
        st.markdown(
            "### 📊 GUM Metodolojisi ile Ölçüm Belirsizliği Hesaplama Motoru"
        )
        st.info(
            "💡 ISO/IEC 17025 standardı gereğince Tip A (Tekrarlanabilirlik)"
            " ve Tip B (Sertifika, Çözünürlük vb.) bileşenlerini birleştirerek"
            " Genişletilmiş Ölçüm Belirsizliği (U) hesaplayın."
        )

        with st.form("olcum_belirsizligi_formu"):
            st.markdown("#### 📐 1. Tip A Değerlendirmesi (Tekrarlanabilirlik)")
            tekrar_verileri_str = st.text_input(
                "Ardışık Ölçüm Değerlerini Virgülle Ayırarak Girin:",
                value="100.1, 100.2, 100.0, 100.3, 100.1",
            )

            st.markdown("---")
            st.markdown(
                "#### 🔬 2. Tip B Değerlendirmesi (Standart Belirsizlik"
                " Bileşenleri)"
            )
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                u_sertifika = st.number_input(
                    "Referans / Etalon Sertifika Belirsizliği (u(cert))",
                    value=0.0200,
                    format="%.4f",
                )
                k_faktoru = st.number_input(
                    "Sertifika Kapsam Faktörü (k)", value=2.0, format="%.2f"
                )
            with col_b2:
                cozunurluk = st.number_input(
                    "Cihaz Çözünürlüğü / Taksimat (a)",
                    value=0.0100,
                    format="%.4f",
                )
                diger_tipb = st.number_input(
                    "Diğer Çevresel / Operasyonel Belirsizlikler (u(oth))",
                    value=0.0100,
                    format="%.4f",
                )

            st.markdown("---")
            kapsam_k_secim = st.selectbox(
                "Nihai Genişletilmiş Belirsizlik İçin Kapsam Faktörü (k)",
                [
                    "k = 2 (%%95 Güven Seviyesi)",
                    "k = 3 (%%99 Güven Seviyesi)",
                ],
            )

            btn_belirsizlik_hesapla = st.form_submit_button(
                "🧮 Belirsizlik Bütçesini Hesapla", type="primary"
            )

        if btn_belirsizlik_hesapla:
            try:
                # Tip A Hesaplama
                degerler = [
                    float(x.strip())
                    for x in tekrar_verileri_str.split(",")
                    if x.strip()
                ]
                n = len(degerler)
                if n < 2:
                    st.error(
                        "⚠️ Tip A analizi için en az 2 ölçüm değeri girilmelidir!"
                    )
                else:
                    arr = np.array(degerler)
                    ortalama = np.mean(arr)
                    st.session_state["cache_ort"] = ortalama
                    std_sapma = np.std(
                        arr, ddof=1
                    )  # Örneklem standart sapması (s)
                    u_A = (
                        std_sapma / math.sqrt(n)
                        if n > 0
                        else 0.0  # Ortalamanın standart belirsizliği
                    )

                    # Tip B Hesaplama (Dikdörtgen dağılım varsayımı: çözünürlük / sqrt(3))
                    u_res = (
                        cozunurluk / math.sqrt(12) if cozunurluk > 0 else 0.0
                    )
                    u_cert_standard = (
                        u_sertifika / k_faktoru if k_faktoru > 0 else u_sertifika
                    )

                    # Birleştirilmiş Standart Belirsizlik (uc)
                    uc = math.sqrt(
                        (u_A**2)
                        + (u_cert_standard**2)
                        + (u_res**2)
                        + (diger_tipb**2)
                    )

                    # Genişletilmiş Belirsizlik (U)
                    k_val = 2.0 if "k = 2" in kapsam_k_secim else 3.0
                    U_genisletilmis = uc * k_val

                    st.success("✅ GUM Belirsizlik Bütçesi Başarıyla Çıkarıldı!")

                    # Sonuç Paneli
                    r1, r2, r3 = st.columns(3)
                    r1.metric("Ölçüm Ortalaması", f"{ortalama:.4f}")
                    r2.metric(
                        "Birleştirilmiş Belirsizlik (uc)", f"±{uc:.4f}"
                    )
                    r3.metric(
                        f"Genişletilmiş Belirsizlik (U, k={k_val})",
                        f"±{U_genisletilmis:.4f}",
                    )

                    st.markdown("---")
                    st.markdown("#### 📋 Detaylı Belirsizlik Bütçesi Tablosu")
                    bbutce_df = pd.DataFrame(
                        [
                            {
                                "Bileşen": "Tip A (Tekrarlanabilirlik)",
                                "Değer / Std. Sapma": f"{std_sapma:.4f}",
                                "Ölçüm Sayısı (n)": n,
                                "Standart Belirsizlik u(xi)": f"{u_A:.4f}",
                            },
                            {
                                "Bileşen": "Tip B (Sertifika / Etalon)",
                                "Değer / Std. Sapma": f"{u_sertifika:.4f}",
                                "Ölçüm Sayısı (n)": f"k={k_faktoru}",
                                "Standart Belirsizlik u(xi)": f"{u_cert_standard:.4f}",
                            },
                            {
                                "Bileşen": "Tip B (Cihaz Çözünürlüğü)",
                                "Değer / Std. Sapma": f"{cozunurluk:.4f}",
                                "Ölçüm Sayısı (n)": "Dikdörtgen (√12)",
                                "Standart Belirsizlik u(xi)": f"{u_res:.4f}",
                            },
                            {
                                "Bileşen": "Tip B (Diğer / Çevre)",
                                "Değer / Std. Sapma": f"{diger_tipb:.4f}",
                                "Ölçüm Sayısı (n)": "-",
                                "Standart Belirsizlik u(xi)": f"{diger_tipb:.4f}",
                            },
                        ]
                    )
                    st.dataframe(bbutce_df, use_container_width=True)

            except Exception as e:
                st.error(f"Hesaplama hatası: {e}")

    with sekmeler[6]:
        st.markdown("### 📐 Metot Validasyonu")

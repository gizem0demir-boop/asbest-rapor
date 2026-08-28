from datetime import datetime, timedelta
import io
import os
from docxtpl import DocxTemplate
import numpy as np
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

    with sekmeler[0]:
        st.markdown("### 📄 FR.71.01.01 Talep ve Teklif Formları Yönetimi")
        teklif_excel = st.file_uploader(
            "📁 Asbest Tutanak Excel Dosyasını Yükleyin (.xlsx)",
            type=["xlsx"],
            key="asbest_tutanak_net_input_v22",
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

        son_dort = (
            st.session_state["teklif_no_val"].split("-")[-1]
            if "-" in st.session_state["teklif_no_val"]
            else "5110"
        )

        with st.form("teklif_formu_net_alan_v22"):
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
            "teklif_net_belge_hazir_v22", False
        ):
            st.session_state["teklif_net_belge_hazir_v22"] = True
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
        sozlesme_excel = st.file_uploader(
            "📁 Sözleşme/Sipariş için Excel Yükleyin (.xlsx)",
            type=["xlsx"],
            key="sozlesme_excel_input_v15",
        )

        soz_firma = st.session_state["firma_val"]
        soz_tarih = st.session_state["tarih_val"]
        soz_no = f"S-{son_dort}"
        soz_adres = st.session_state["adres_val"]
        soz_tel = st.session_state["tel_val"]

        with st.form("sozlesme_formu_alan_v15"):
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
                        "imza_durumu": secilen_durum,
                    }
                )
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

    with sekmeler[2]:
        st.markdown(
            "### 📝 Saha Kayıtları: KKD ve Asbest Risk Değerlendirmesi"
        )
        with st.form("kkd_formu_hazirla_v15"):
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
            btn_kkd_indir = st.form_submit_button(
                "📥 KKD Formunu Doldur ve İndir", type="primary"
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
                st.success("✅ KKD formu hazırlandı!")
                st.download_button(
                    label="⬇️ KKD Formunu İndir (.docx)",
                    data=output_kkd.getvalue(),
                    file_name=f"KKD_Formu_{kkd_teklif_no}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                )

    with sekmeler[3]:
        st.markdown(
            "### 📅 ISO/IEC 17025 Cihaz Kalibrasyon ve Periyodik Kontrol"
            " Takip Paneli"
        )
        st.info(
            "💡 Hizmet dışı/arızalı cihazlar hariç tutularak aktif kullanımda"
            " olan cihazlar listelenmektedir."
        )

        cihaz_excel = st.file_uploader(
            "📁 Farklı Bir Cihaz Envanteri Yüklemek İçin (.xlsx)",
            type=["xlsx"],
            key="cihaz_envanter_excel_input_v3",
        )

        hedef_dosya = (
            cihaz_excel
            if cihaz_excel is not None
            else "LS.66.03.07 Kalibrasyon Takip ve Cihaz Listesi.xlsx -10-20.07.2026.xlsx"
        )

        try:
            if os.path.exists(hedef_dosya) or cihaz_excel is not None:
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

                            kalibrasyon_durumu = "Kalibrasyon Gerekmez / Belirsiz"
                            kalan_gun = 9999

                            import re

                            tarih_eslesmeleri = re.findall(
                                r"\d{4}[-/.]\d{2}[-/.]\d{2}|\d{2}[-/.]\d{2}[-/.]\d{4}",
                                tarih_hucre,
                            )
                            if tarih_eslesmeleri:
                                t_str = tarih_eslesmeleri[-1]
                                try:
                                    for fmt in (
                                        "%Y-%m-%d",
                                        "%d.%m.%Y",
                                        "%d/%m/%Y",
                                    ):
                                        try:
                                            dt_obj = datetime.strptime(
                                                t_str.replace("/", "."), fmt
                                            )
                                            kalan_gun = (
                                                dt_obj - bugun
                                            ).days
                                            break
                                        except:
                                            pass
                                except:
                                    pass

                            if "gerekmez" in tarih_hucre.lower():
                                kalibrasyon_durumu = (
                                    "🟢 Kalibrasyon Gerekmez (Ara Kontrol)"
                                )
                            elif kalan_gun == 9999:
                                kalibrasyon_durumu = "⚪ Takip Edilmiyor / Diğer"
                            elif kalan_gun < 0:
                                kalibrasyon_durumu = (
                                    f"🔴 SÜRESİ GEÇTİ! ({abs(kalan_gun)} gün önce)"
                                )
                            elif kalan_gun <= 30:
                                kalibrasyon_durumu = (
                                    f"🟡 SÜRESİ YAKLAŞIYOR ({kalan_gun} gün kaldı)"
                                )
                            else:
                                kalibrasyon_durumu = (
                                    f"🟢 Güncel ({kalan_gun} gün kaldı)"
                                )

                            tum_cihazlar.append(
                                {
                                    "No": int(float(val)),
                                    "Cihaz Adı / Marka": str(row.iloc[1]).strip()
                                    if pd.notna(row.iloc[1])
                                    else "-",
                                    "Seri No": str(row.iloc[2]).strip()
                                    if pd.notna(row.iloc[2])
                                    else "-",
                                    "Parametre / Metot": str(row.iloc[3]).strip()
                                    if len(row) > 3 and pd.notna(row.iloc[3])
                                    else "-",
                                    "Kalibrasyon Bilgisi": tarih_hucre,
                                    "Durum": kalibrasyon_durumu,
                                }
                            )

                df_envanter = pd.DataFrame(tum_cihazlar)

                c1, c2, c3 = st.columns(3)
                c1.metric("Aktif Cihaz Sayısı", len(df_envanter))
                suresi_gelen = len(
                    df_envanter[
                        df_envanter["Durum"].str.contains(
                            "YAKLAŞIYOR", case=False
                        )
                    ]
                )
                suresi_gecen = len(
                    df_envanter[
                        df_envanter["Durum"].str.contains("GEÇTİ", case=False)
                    ]
                )
                c2.metric("Süresi Yaklaşan (<30 Gün)", suresi_gelen)
                c3.metric("Süresi Geçen", suresi_gecen)

                arama = st.text_input(
                    "🔍 Aktif Cihaz Listesinde Ara (Marka, Seri No veya Ad):"
                )
                if arama:
                    df_envanter = df_envanter[
                        df_envanter.apply(
                            lambda row: row.astype(str)
                            .str.contains(arama, case=False)
                            .any(),
                            axis=1,
                        )
                    ]

                st.dataframe(df_envanter, use_container_width=True)
            else:
                st.warning("⚠️ Excel dosyası bulunamadı.")
        except Exception as e:
            st.error(f"Dosya işlenirken hata oluştu: {e}")

    with sekmeler[4]:
        st.markdown(
            "### ⚖️ Kalibrasyon Kabul ve Uygunluk Analizi (Kriter Bazlı)"
        )
        st.info(
            "💡 Envanter dosyasındaki cihazlara ait kabul kriterleri otomatik"
            " olarak yüklenmiştir. Cihazınızı seçerek uygunluk değerlendirmesi"
            " yapabilirsiniz."
        )

        try:
            if os.path.exists(hedef_dosya) or cihaz_excel is not None:
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
                            c_kriter = (
                                str(row.iloc[9]).strip()
                                if len(row) > 9 and pd.notna(row.iloc[9])
                                else "--"
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
                        "Değerlendirilecek Cihazı Seçin:",
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
                        f"📋 **Cihazın Envanterdeki Kabul Kriteri:**\n\n`{secilen_cihaz['kriter']}`"
                    )

                    with st.form("kabul_degerlendirme_formu"):
                        col_m1, col_m2 = st.columns(2)
                        with col_m1:
                            ref_deger = st.number_input(
                                "Referans Değer (Etalon)", value=100.0
                            )
                        with col_m2:
                            olculen_deger = st.number_input(
                                "Cihazın Ölçülen Değeri", value=100.2
                            )

                        maks_tolerans = st.number_input(
                            "Maksimum İzin Verilen Tolerans / Hata Sınırı"
                            " (±)",
                            value=1.0,
                        )
                        btn_kabul_yap = st.form_submit_button(
                            "🔍 Uygunluk Değerlendir", type="primary"
                        )

                    if btn_kabul_yap:
                        fark = abs(olculen_deger - ref_deger)
                        if fark <= maks_tolerans:
                            st.success(
                                f"✅ **KABUL (UYGUN)**: Ölçülen sapma"
                                f" ({fark:.4f}), tanımlanan tolerans sınırları"
                                f" (±{maks_tolerans}) içindedir."
                            )
                        else:
                            st.error(
                                f"❌ **RED (UYGUN DEĞİL)**: Ölçülen sapma"
                                f" ({fark:.4f}), tolerans sınırını"
                                f" (±{maks_tolerans}) aşıyor!"
                            )
                else:
                    st.warning("Uygun cihaz bulunamadı.")
        except Exception as e:
            st.error(f"Kabul kriterleri yüklenirken hata oluştu: {e}")

    with sekmeler[5]:
        st.markdown("### 📊 Ölçüm Belirsizliği Hesaplamaları")
    with sekmeler[6]:
        st.markdown("### 📐 Metot Validasyonu")

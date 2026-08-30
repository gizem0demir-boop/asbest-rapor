import datetime
import os
import re
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
import logging

def render():
    st.title("🏗️ Yıkım Planı ve Yasal Evrak Modülü")
    st.markdown("---")

    df_muellif, df_muteahhit = veritabani_yukle()
    if df_muellif.empty or df_muteahhit.empty:
        st.warning(f"⚠️ '{EXCEL_VT_YOLU}' dosyasından veriler okunamadı. Lütfen kontrol edin.")
        return

    # 1. ADIM: DOSYAYI EN BAŞTA TEK BİR YERDE YÜKLET VE HAFIZAYA AL
    st.subheader("📂 1. Adım: Yapı Bilgi Tutanak / Belge Yükleme")
    tutanak_file = st.file_uploader("Yapı Bilgilerini İçeren Excel Dosyasını Yükleyin:", type=SUPPORTED_FILE_TYPES, key="ana_tutanak_dosyasi")
    
    # Dosya ilk kez yüklendiğinde ya da değiştiğinde okuyup session_state'e atalım
    if tutanak_file is not None:
        if "son_okunan_dosya" not in st.session_state or st.session_state.get("son_okunan_dosya") != tutanak_file.name:
            st.session_state["son_okunan_dosya"] = tutanak_file.name
            st.session_state["yapi_bilgileri"] = read_fenni_mesul_details(tutanak_file)
            st.success("✅ Tutanak başarıyla okundu ve hafızaya alındı!")
    else:
        # Dosya yoksa boş standart değerler atayalım
        if "yapi_bilgileri" not in st.session_state:
            st.session_state["yapi_bilgileri"] = {"yapi_adresi": "-", "ada_parsel": "-", "il_ilce": "-", "idare": "-"}

    # Hafızadaki verileri pratik değişkenlere alalım
    aktif_bilgi = st.session_state["yapi_bilgileri"]

    st.markdown("---")

    # 2. ADIM: EVRAK TÜRÜ SEÇİMİ
    alt_islem = st.selectbox(
        "📌 2. Adım: Oluşturulacak Evrak Türünü Seçin:",
        [
            "-- Seçiniz --",
            "🤝 Müellif - Müteahhit Yıkım Sözleşmesi",
            "📜 Fenni Mesul Taahhütnamesi",
            "📝 Müellif Taahhütnamesi (İdareye Verilecek - Form 2)",
            "🏗️ Yıkım Planı Raporu",
        ],
    )
    st.markdown("---")

    # Ortak adres ve ada/parsel alanlarını hafızadan otomatik doldurarak gösterelim
    if alt_islem != "-- Seçiniz --":
        st.info(f"💡 Hafızadaki Yapı Bilgileri -> Adres: **{aktif_bilgi.get('yapi_adresi')}** | Ada/Parsel: **{aktif_bilgi.get('ada_parsel')}**")

    # 1. MÜELLİF - MÜTEAHHİT SÖZLEŞMESİ
    if alt_islem == "🤝 Müellif - Müteahhit Yıkım Sözleşmesi":
        st.subheader("🤝 Müellif ve Müteahhit Yıkım Sözleşmesi")
        col1, col2 = st.columns(2)
        with col1:
            secilen_muellif_ad = st.selectbox("Müellif Seçiniz:", df_muellif["Ad_Soyad"].tolist(), key="soz_mue_secim")
            m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_muellif_ad].iloc[0]
            st.text_input("Oda Sicil No:", value=str(m_satir.get("Oda_Sicil_No", "")), disabled=True)
            st.text_input("TC Kimlik No:", value=str(m_satir.get("TC_No", "")), disabled=True)
        with col2:
            secilen_mut_firma = st.selectbox("Müteahhit Firma Seçiniz:", df_muteahhit["Firma_Unvani"].tolist(), key="soz_mut_secim")
            mut_satir = df_muteahhit[df_muteahhit["Firma_Unvani"] == secilen_mut_firma].iloc[0]
            st.text_input("Yetkili Ad Soyad:", value=str(mut_satir.get("Yetkili_Ad_Soyad", "")), disabled=True)
            st.text_input("Vergi No / TC:", value=str(mut_satir.get("Vergi_No_TC", "")), disabled=True)

        col3, col4 = st.columns(2)
        yapi_adresi = col3.text_input("Yapı Adresi:", value=aktif_bilgi.get("yapi_adresi", ""), key="soz_adres")
        ada_parsel = col4.text_input("Ada / Parsel:", value=aktif_bilgi.get("ada_parsel", ""), key="soz_ada")

        sozlesme_suresi = st.number_input("Sözleşme Süresi (Gün):", value=90, key="soz_sure")
        ucret = st.text_input("Anlaşma Ücreti (TL):", value="1500 TL + KDV", key="soz_ucret")

        if st.button("🚀 Sözleşmeyi Oluştur ve İndir", type="primary", key="btn_soz"):
            context = {
                "muellif_adi": m_satir.get("Ad_Soyad"), "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                "muellif_tc": m_satir.get("TC_No"), "muellif_tel": m_satir.get("Telefon"),
                "muteahhit_firma": mut_satir.get("Firma_Unvani"), "muteahhit_yetkili": mut_satir.get("Yetkili_Ad_Soyad"),
                "muteahhit_vno": mut_satir.get("Vergi_No_TC"), "muteahhit_adres": mut_satir.get("Adres"),
                "muteahhit_tel": mut_satir.get("Telefon"), "yapi_adresi": yapi_adresi, "ada_parsel": ada_parsel,
                "sure": sozlesme_suresi, "ucret": ucret, "ucret_yazi": sayiyi_yaziya_cevir(ucret),
                "tarih": datetime.date.today().strftime("%d.%m.%Y")
            }
            sablon_yolu = "templates/yikim_sozlesme_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis = "Yikim_Sozlesmesi.docx"
                doc.save(cikis)
                with open(cikis, "rb") as f:
                    st.download_button("📥 Sözleşmeyi İndir", f, file_name="Yikim_Sozlesmesi.docx", key="dl_soz")
                st.success("✅ Sözleşme başarıyla oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: {sablon_yolu}")

    # 2. FENNİ MESUL TAAHHÜTNAMESİ
    elif alt_islem == "📜 Fenni Mesul Taahhütnamesi":
        st.subheader("📜 Fenni Mesul Taahhütnamesi Hazırlama")
        secilen_fenni = st.selectbox("Fenni Mesul Seçin:", df_muellif["Ad_Soyad"].tolist(), key="fenni_secim")
        f_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_fenni].iloc[0]

        col1, col2 = st.columns(2)
        yapi_adresi = col1.text_input("Yapı Adresi:", value=aktif_bilgi.get("yapi_adresi", "-"), key="fenni_adres")
        ada_parsel = col2.text_input("Ada / Parsel:", value=aktif_bilgi.get("ada_parsel", "-"), key="fenni_ada")

        if st.button("🚀 Taahhütnameyi Oluştur", type="primary", key="btn_fenni"):
            context = {
                "fenni_adi": f_satir.get("Ad_Soyad"), "fenni_tc": f_satir.get("TC_No"),
                "fenni_oda_no": f_satir.get("Oda_Sicil_No"), "yapi_adresi": yapi_adresi,
                "ada_parsel": ada_parsel, "tarih": datetime.date.today().strftime("%d.%m.%Y")
            }
            sablon_yolu = "templates/fenni_mesul_taahhutname_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis = "Fenni_Mesul_Taahhutnamesi.docx"
                doc.save(cikis)
                with open(cikis, "rb") as f:
                    st.download_button("📥 Taahhütnameyi İndir", f, file_name="Fenni_Mesul_Taahhutnamesi.docx", key="dl_fenni")
                st.success("✅ Fenni Mesul Taahhütnamesi oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: {sablon_yolu}")

    # 3. MÜELLİF TAAHHÜTNAMESİ (FORM 2)
    elif alt_islem == "📝 Müellif Taahhütnamesi (İdareye Verilecek - Form 2)":
        st.subheader("📝 Müellif Taahhütnamesi (Form 2)")
        secilen_mue = st.selectbox("Müellif Seçin:", df_muellif["Ad_Soyad"].tolist(), key="form2_mue")
        m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_mue].iloc[0]
        idare_adi = st.text_input("İlgili İdare / Belediye:", value=aktif_bilgi.get("idare", "Belediye Başkanlığı"), key="form2_idare")

        col1, col2 = st.columns(2)
        yapi_adresi = col1.text_input("Yapı Adresi:", value=aktif_bilgi.get("yapi_adresi", "-"), key="form2_adres")
        ada_parsel = col2.text_input("Ada / Parsel:", value=aktif_bilgi.get("ada_parsel", "-"), key="form2_ada")

        if st.button("🚀 Form 2 Oluştur", type="primary", key="btn_form2"):
            context = {
                "idare_adi": idare_adi, "muellif_adi": m_satir.get("Ad_Soyad"),
                "muellif_tc": m_satir.get("TC_No"), "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                "yapi_adresi": yapi_adresi, "ada_parsel": ada_parsel,
                "tarih": datetime.date.today().strftime("%d.%m.%Y")
            }
            sablon_yolu = "templates/form2_taahhutname_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis = "Form2_Muellif_Taahhutnamesi.docx"
                doc.save(cikis)
                with open(cikis, "rb") as f:
                    st.download_button("📥 Form 2 İndir", f, file_name="Form2_Muellif_Taahhutnamesi.docx", key="dl_form2")
                st.success("✅ Form 2 Taahhütnamesi oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: {sablon_yolu}")

    # 4. YIKIM PLANI RAPORU
    elif alt_islem == "🏗️ Yıkım Planı Raporu":
        st.subheader("🏗️ Yıkım Planı Raporu Oluşturucu")
        col_mue, col_mut = st.columns(2)
        with col_mue:
            secilen_mue = st.selectbox("Proje Müellifi Seçin:", df_muellif["Ad_Soyad"].tolist(), key="yp_mue")
            m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_mue].iloc[0]
        with col_mut:
            secilen_mut = st.selectbox("Müteahhit Firma Seçin:", df_muteahhit["Firma_Unvani"].tolist(), key="yp_mut")
            mut_satir = df_muteahhit[df_muteahhit["Firma_Unvani"] == secilen_mut].iloc[0]

        col1, col2 = st.columns(2)
        yapi_adresi = col1.text_input("Yapı Adresi:", value=aktif_bilgi.get("yapi_adresi", ""), key="yp_adres")
        ada_parsel = col2.text_input("Ada / Parsel:", value=aktif_bilgi.get("ada_parsel", ""), key="yp_ada")

        col3, col4 = st.columns(2)
        yikim_yontemi = col3.selectbox("Yıkım Yöntemi:", ["Mekanik Yıkım (Ekskavatör)", "Kademeli Yıkım", "Elle + Mekanik Yıkım"], key="yp_yontem")
        muhit = col4.selectbox("Saha Konumu:", ["Meskun Mahal", "Sanayi Bölgesi", "Açık / Kırsal"], key="yp_muhit")

        if st.button("🚀 Yıkım Planı Raporunu Oluştur", type="primary", key="btn_yp"):
            context = {
                "muellif_adi": m_satir.get("Ad_Soyad"), "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                "muellif_tc": m_satir.get("TC_No"), "muellif_tel": m_satir.get("Telefon"),
                "muteahhit_firma": mut_satir.get("Firma_Unvani"), "muteahhit_yetkili": mut_satir.get("Yetkili_Ad_Soyad"),
                "muteahhit_vno": mut_satir.get("Vergi_No_TC"), "muteahhit_adres": mut_satir.get("Adres"),
                "muteahhit_tel": mut_satir.get("Telefon"), "yapi_adresi": yapi_adresi, "ada_parsel": ada_parsel,
                "yikim_yontemi": yikim_yontemi, "muhit": muhit, "tarih": datetime.date.today().strftime("%d.%m.%Y")
            }
            sablon_yolu = "templates/yikim_plani_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis = "Yikim_Plani_Raporu.docx"
                doc.save(cikis)
                with open(cikis, "rb") as f:
                    st.download_button("📥 Raporu İndir", f, file_name="Yikim_Plani_Raporu.docx", key="dl_yp")
                st.success("✅ Yıkım Planı Raporu başarıyla oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: {sablon_yolu}")

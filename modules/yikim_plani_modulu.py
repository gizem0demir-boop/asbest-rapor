import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from utils.excel_parser import read_tutanak_details
import os
import datetime
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from utils.excel_parser import read_tutanak_details

EXCEL_VT_YOLU = "veritabani.xlsx"

@st.cache_data(ttl=60)
def veritabani_yukle():
    if not os.path.exists(EXCEL_VT_YOLU):
        st.error(f"❌ '{EXCEL_VT_YOLU}' bulunamadı. Lütfen repo dizinine Excel dosyasını ekleyin.")
        return pd.DataFrame(), pd.DataFrame()

    try:
        df_muellif = pd.read_excel(EXCEL_VT_YOLU, sheet_name=0)
        df_muteahhit = pd.read_excel(EXCEL_VT_YOLU, sheet_name=1)
        return df_muellif, df_muteahhit
    except Exception as e:
        st.error(f"Excel okunurken hata oluştu: {e}")
        return pd.DataFrame(), pd.DataFrame()

def render():
    st.title("🏗️ Yıkım Planı ve Yasal Evrak Modülü")
    st.caption("Dinamik Veritabanı Entegreli Arayüz")
    st.markdown("---")

    df_muellif, df_muteahhit = veritabani_yukle()

    if df_muellif.empty or df_muteahhit.empty:
        st.warning("⚠️ Excel veritabanından veri okunamadı. Lütfen 'veritabani.xlsx' dosyasını kontrol edin.")
        return

    alt_islem = st.selectbox(
        "📌 Oluşturulacak Evrak Türünü Seçin:",
        [
            "-- Seçiniz --",
            "🤝 Müellif - Müteahhit Yıkım Sözleşmesi",
            "📜 Fenni Mesul Taahhütnamesi",
            "📝 Müellif Taahhütnamesi (İdareye Verilecek - Form 2)",
            "🏗️ Yıkım Planı Raporu",
        ],
    )

    st.markdown("---")

    # 1. MÜELLİF - MÜTEAHHİT SÖZLEŞMESİ
    if alt_islem == "🤝 Müellif - Müteahhit Yıkım Sözleşmesi":
        st.subheader("🤝 Müellif ve Müteahhit Yıkım Sözleşmesi")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### 👷 Proje Müellifi (Mühendis)")
            secilen_muellif_ad = st.selectbox("Müellif Seçiniz:", df_muellif["Ad_Soyad"].tolist(), key="soz_mue_secim")
            m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_muellif_ad].iloc[0]

            st.text_input("Oda Sicil No:", value=str(m_satir.get("Oda_Sicil_No", "")), disabled=True)
            st.text_input("TC Kimlik No:", value=str(m_satir.get("TC_No", "")), disabled=True)
            st.text_input("Müellif Tel:", value=str(m_satir.get("Telefon", "")), disabled=True)

        with col2:
            st.markdown("### 🏢 Müteahhit / İşveren")
            secilen_mut_firma = st.selectbox("Müteahhit Firma Seçiniz:", df_muteahhit["Firma_Unvani"].tolist(), key="soz_mut_secim")
            mut_satir = df_muteahhit[df_muteahhit["Firma_Unvani"] == secilen_mut_firma].iloc[0]

            st.text_input("Yetkili Ad Soyad:", value=str(mut_satir.get("Yetkili_Ad_Soyad", "")), disabled=True)
            st.text_input("Vergi No / TC:", value=str(mut_satir.get("Vergi_No_TC", "")), disabled=True)
            st.text_input("Firma Tel:", value=str(mut_satir.get("Telefon", "")), disabled=True)

        st.markdown("### 🗺️ Yapı ve Saha Bilgileri")
        tutanak_file = st.file_uploader("📂 Yapı Bilgi Tutanak Excel'ini Yükleyin (Ada/Parsel/Adres için):", type=["xlsx", "xls"], key="soz_tutanak")

        col3, col4 = st.columns(2)
        if tutanak_file:
            yapi_data = read_tutanak_details(tutanak_file)
            yapi_adresi = col3.text_input("Yapı Adresi:", value=yapi_data.get("yapi_adresi", ""))
            ada_parsel = col4.text_input("Ada / Parsel:", value=yapi_data.get("ada_parsel", ""))
        else:
            yapi_adresi = col3.text_input("Yapı Adresi:", value="Kazım Karabekir Mah. 220. Sok. No: 78 Bağcılar, İstanbul")
            ada_parsel = col4.text_input("Ada / Parsel:", value="853 Ada 20 Parsel")

        sozlesme_suresi = st.number_input("Sözleşme Süresi (Gün):", value=90, step=15)
        ucret = st.text_input("Anlaşma Ücreti (TL):", value="1500 TL + KDV")

        if st.button("🚀 Yıkım Sözleşmesini Doldur ve İndir", type="primary"):
            context = {
                "muellif_adi": m_satir.get("Ad_Soyad"),
                "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                "muellif_tc": m_satir.get("TC_No"),
                "muellif_tel": m_satir.get("Telefon"),
                "muteahhit_firma": mut_satir.get("Firma_Unvani"),
                "muteahhit_yetkili": mut_satir.get("Yetkili_Ad_Soyad"),
                "muteahhit_vno": mut_satir.get("Vergi_No_TC"),
                "muteahhit_adres": mut_satir.get("Adres"),
                "muteahhit_tel": mut_satir.get("Telefon"),
                "yapi_adresi": yapi_adresi,
                "ada_parsel": ada_parsel,
                "sure": sozlesme_suresi,
                "ucret": ucret,
                "tarih": datetime.date.today().strftime("%d.%m.%Y")
            }

            sablon_yolu = "templates/yikim_sozlesme_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis_yolu = "Yikim_Sozlesmesi_Cikti.docx"
                doc.save(cikis_yolu)

                with open(cikis_yolu, "rb") as f:
                    st.download_button("📥 Hazır Sözleşmeyi İndir", f, file_name="Yikim_Sozlesmesi.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                st.success("✅ Yıkım Sözleşmesi başarıyla oluşturuldu!")
            else:
                st.error(f"❌ Şablon dosyası bulunamadı: '{sablon_yolu}'")

    # 2. FENNİ MESUL TAAHHÜTNAMESİ
    elif alt_islem == "📜 Fenni Mesul Taahhütnamesi":
        st.subheader("📜 Fenni Mesul Taahhütnamesi Hazırlama")

        secilen_fenni = st.selectbox("Fenni Mesul (Mühendis) Seçin:", df_muellif["Ad_Soyad"].tolist(), key="fenni_secim")
        f_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_fenni].iloc[0]

        st.info(f"Seçilen Fenni Mesul: **{f_satir.get('Ad_Soyad')}** | Oda No: **{f_satir.get('Oda_Sicil_No')}** | TC: **{f_satir.get('TC_No')}**")

        tutanak_file = st.file_uploader("📂 Tutanak Excel Yükleyin (İsteğe Bağlı):", type=["xlsx", "xls"], key="fenni_tutanak")
        col1, col2 = st.columns(2)
        if tutanak_file:
            yapi_data = read_tutanak_details(tutanak_file)
            yapi_adresi = col1.text_input("Yapı Adresi:", value=yapi_data.get("yapi_adresi", ""), key="fenni_adres")
            ada_parsel = col2.text_input("Ada / Parsel:", value=yapi_data.get("ada_parsel", ""), key="fenni_ada")
        else:
            yapi_adresi = col1.text_input("Yapı Adresi:", value="-", key="fenni_adres")
            ada_parsel = col2.text_input("Ada / Parsel:", value="-", key="fenni_ada")

        if st.button("🚀 Fenni Mesul Taahhütnamesi Oluştur", type="primary"):
            context = {
                "fenni_adi": f_satir.get("Ad_Soyad"),
                "fenni_tc": f_satir.get("TC_No"),
                "fenni_oda_no": f_satir.get("Oda_Sicil_No"),
                "yapi_adresi": yapi_adresi,
                "ada_parsel": ada_parsel,
                "tarih": datetime.date.today().strftime("%d.%m.%Y")
            }
            sablon_yolu = "templates/fenni_mesul_taahhutname_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis = "Fenni_Mesul_Taahhutnamesi.docx"
                doc.save(cikis)
                with open(cikis, "rb") as f:
                    st.download_button("📥 Taahhütnameyi İndir", f, file_name="Fenni_Mesul_Taahhutnamesi.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                st.success("✅ Fenni Mesul Taahhütnamesi oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: '{sablon_yolu}'")

    # 3. MÜELLİF TAAHHÜTNAMESİ (FORM 2)
    elif alt_islem == "📝 Müellif Taahhütnamesi (İdareye Verilecek - Form 2)":
        st.subheader("📝 Müellif Taahhütnamesi (Form 2)")

        secilen_mue = st.selectbox("Müellif Seçin:", df_muellif["Ad_Soyad"].tolist(), key="form2_mue")
        m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_mue].iloc[0]

        idare_adi = st.text_input("İlgili İdare / Belediye Adı:", value="Kadıköy Belediye Başkanlığı Yapı Kontrol Müdürlüğü'ne")

        tutanak_file = st.file_uploader("📂 Tutanak Excel Yükleyin:", type=["xlsx", "xls"], key="form2_tutanak")
        col1, col2 = st.columns(2)
        if tutanak_file:
            yapi_data = read_tutanak_details(tutanak_file)
            yapi_adresi = col1.text_input("Yapı Adresi:", value=yapi_data.get("yapi_adresi", ""), key="form2_adres")
            ada_parsel = col2.text_input("Ada / Parsel:", value=yapi_data.get("ada_parsel", ""), key="form2_ada")
        else:
            yapi_adresi = col1.text_input("Yapı Adresi:", value="-", key="form2_adres")
            ada_parsel = col2.text_input("Ada / Parsel:", value="-", key="form2_ada")

        if st.button("🚀 Form 2 Taahhütnamesi Oluştur", type="primary"):
            context = {
                "idare_adi": idare_adi,
                "muellif_adi": m_satir.get("Ad_Soyad"),
                "muellif_tc": m_satir.get("TC_No"),
                "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                "yapi_adresi": yapi_adresi,
                "ada_parsel": ada_parsel,
                "tarih": datetime.date.today().strftime("%d.%m.%Y")
            }
            sablon_yolu = "templates/form2_taahhutname_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis = "Form2_Muellif_Taahhutnamesi.docx"
                doc.save(cikis)
                with open(cikis, "rb") as f:
                    st.download_button("📥 Form 2 İndir", f, file_name="Form2_Muellif_Taahhutnamesi.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                st.success("✅ Form 2 Taahhütnamesi oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: '{sablon_yolu}'")

    # 4. YIKIM PLANI RAPORU
    elif alt_islem == "🏗️ Yıkım Planı Raporu":
        st.subheader("🏗️ Yıkım Planı Raporu Oluşturucu")

        secilen_mue = st.selectbox("Proje Müellifi Seçin:", df_muellif["Ad_Soyad"].tolist(), key="yp_mue")
        m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_mue].iloc[0]

        secilen_mut = st.selectbox("Müteahhit Firma Seçin:", df_muteahhit["Firma_Unvani"].tolist(), key="yp_mut")
        mut_satir = df_muteahhit[df_muteahhit["Firma_Unvani"] == secilen_mut].iloc[0]

        tutanak_file = st.file_uploader("📂 Tutanak Excel Yükleyin:", type=["xlsx", "xls"], key="yp_tutanak")
        col1, col2 = st.columns(2)
        if tutanak_file:
            yapi_data = read_tutanak_details(tutanak_file)
            yapi_adresi = col1.text_input("Yapı Adresi:", value=yapi_data.get("yapi_adresi", ""), key="yp_adres")
            ada_parsel = col2.text_input("Ada / Parsel:", value=yapi_data.get("ada_parsel", ""), key="yp_ada")
        else:
            yapi_adresi = col1.text_input("Yapı Adresi:", value="-", key="yp_adres")
            ada_parsel = col2.text_input("Ada / Parsel:", value="-", key="yp_ada")

        col3, col4 = st.columns(2)
        yikim_yontemi = col3.selectbox("Yıkım Yöntemi:", ["Mekanik Yıkım (Ekskavatör)", "Kademeli Yıkım", "Elle + Mekanik Yıkım"])
        muhit = col4.selectbox("Saha Konumu:", ["Meskun Mahal", "Sanayi Bölgesi", "Açık / Kırsal"])

        if st.button("🚀 Yıkım Planı Raporunu Oluştur", type="primary"):
            context = {
                "muellif_adi": m_satir.get("Ad_Soyad"),
                "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                "muteahhit_firma": mut_satir.get("Firma_Unvani"),
                "yapi_adresi": yapi_adresi,
                "ada_parsel": ada_parsel,
                "yikim_yontemi": yikim_yontemi,
                "muhit": muhit,
                "tarih": datetime.date.today().strftime("%d.%m.%Y")
            }
            sablon_yolu = "templates/yikim_plani_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis = "Yikim_Plani_Raporu.docx"
                doc.save(cikis)
                with open(cikis, "rb") as f:
                    st.download_button("📥 Yıkım Planı Raporunu İndir", f, file_name="Yikim_Plani_Raporu.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                st.success("✅ Yıkım Planı Raporu başarıyla oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: '{sablon_yolu}'")

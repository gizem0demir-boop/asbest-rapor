import datetime
import os
import re
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
import logging

EXCEL_VT_YOLU = "veritabani.xlsx"

def sayiyi_yaziya_cevir(tutar_str):
    """Girilen tutar ifadesindeki sayıları Türkçede yasal evrak formatında yazıya çevirir."""
    try:
        rakamlar = re.findall(r'\d+', str(tutar_str))
        if not rakamlar:
            return tutar_str
        
        tutar = int("".join(rakamlar))
        if tutar == 0:
            return "Sıfır Türk Lirası"

        birler = ["", "Bir", "İki", "Üç", "Dört", "Beş", "Altı", "Yedi", "Sekiz", "Dokuz"]
        onlar = ["", "On", "Yirmi", "Otuz", "Kırk", "Elli", "Altmış", "Yetmiş", "Sekiz", "Doksan"]

        milyon = (tutar // 1_000_000) % 1000
        bin_grubu = (tutar // 1_000) % 1000
        kalan = tutar % 1000

        parcalar = []
        if milyon > 0:
            if milyon == 1:
                parcalar.append("BirMilyon")
            else:
                y = milyon // 100
                o = (milyon % 100) // 10
                b = milyon % 10
                m_str = ""
                if y > 0:
                    m_str += ("BirYüz" if y == 1 else birler[y] + "Yüz")
                if o > 0:
                    m_str += onlar[o]
                if b > 0:
                    m_str += birler[b]
                parcalar.append(m_str + "Milyon")

        if bin_grubu > 0:
            if bin_grubu == 1:
                parcalar.append("Bin")
            else:
                y = bin_grubu // 100
                o = (bin_grubu % 100) // 10
                b = bin_grubu % 10
                b_str = ""
                if y > 0:
                    b_str += ("BirYüz" if y == 1 else birler[y] + "Yüz")
                if o > 0:
                    b_str += onlar[o]
                if b > 0:
                    b_str += birler[b]
                parcalar.append(b_str + "Bin")

        if kalan > 0 or tutar == 0:
            y = kalan // 100
            o = (kalan % 100) // 10
            b = kalan % 10
            k_str = ""
            if y > 0:
                k_str += ("BirYüz" if y == 1 else birler[y] + "Yüz")
            if o > 0:
                k_str += onlar[o]
            if b > 0:
                k_str += birler[b]
            if k_str:
                parcalar.append(k_str)

        tam_metin = "".join(parcalar)
        bosluklu = re.sub(r'(?<!^)(?=[A-Z])', ' ', tam_metin)
        return bosluklu + " Türk Lirası"
    except Exception:
        return tutar_str


def adresinden_il_ilce_bul(adres_metni):
    if not adres_metni:
        return "-", "-"
    parts = [p.strip() for p in adres_metni.split(',')]
    il, ilce = "-", "-"
    if len(parts) >= 2:
        il = parts[-1]
        ilce_aday = parts[-2]
        ilce_parcalari = ilce_aday.split(' ')
        ilce = ilce_parcalari[-1]
    return il, ilce


def read_fenni_mesul_details(tutanak_file):
    info = {
        "yapi_adresi": "-",
        "ada_parsel": "-",
        "il_ilce": "-",
        "idare": "-",
        "yapi_sahibi": "-"
    }
    try:
        if hasattr(tutanak_file, "seek"):
            tutanak_file.seek(0)
            
        xls = pd.ExcelFile(tutanak_file)
        sheet_to_load = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_to_load, header=None)
        
        ada_val, parsel_val, adres_val, sahip_val = "", "", "", ""
        
        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row):
                if pd.isna(val):
                    continue
                val_str = str(val).strip()
                val_lower = val_str.lower()
                
                if "adres" in val_lower or "firma adresi" in val_lower:
                    if ":" in val_str:
                        parcalar = val_str.split(":", 1)
                        if len(parcalar) > 1 and len(parcalar[1].strip()) > 3:
                            adres_val = parcalar[1].strip()
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        adres_val = str(row.iloc[c_idx + 1]).strip()
                
                if "ada" in val_lower:
                    if ":" in val_str:
                        m = re.search(r'(?:ada)[^0-9]*([0-9\w\-]+)', val_str, re.IGNORECASE)
                        if m: 
                            val_bulunan = m.group(1)
                            if val_bulunan.lower() not in ['o', 'yok', '']:
                                ada_val = val_bulunan
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        val_yan = str(row.iloc[c_idx + 1]).strip()
                        if val_yan.lower() not in ['o', 'yok', '-', '']:
                            ada_val = val_yan
                        
                if "parsel" in val_lower:
                    if ":" in val_str:
                        m = re.search(r'(?:parsel)[^0-9]*([0-9\w\-]+)', val_str, re.IGNORECASE)
                        if m: 
                            val_bulunan = m.group(1)
                            if val_bulunan.lower() not in ['yok', '']:
                                parsel_val = val_bulunan
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        val_yan = str(row.iloc[c_idx + 1]).strip()
                        if val_yan.lower() not in ['yok', '-', '']:
                            parsel_val = val_yan

                if "yAPI SAHİBİ" in val_str.upper() or "İŞVEREN" in val_str.upper():
                    if ":" in val_str:
                        parcalar = val_str.split(":", 1)
                        if len(parcalar) > 1 and len(parcalar[1].strip()) > 2:
                            sahip_val = parcalar[1].strip()
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        sahip_val = str(row.iloc[c_idx + 1]).strip()

        if adres_val and adres_val != "-":
            info["yapi_adresi"] = adres_val
            il, ilce = adresinden_il_ilce_bul(adres_val)
            if il != "-" and ilce != "-":
                info["il_ilce"] = f"{il} / {ilce}"
                info["idare"] = f"{ilce} Belediyesi"
            elif ilce != "-":
                info["idare"] = f"{ilce} Belediyesi"
            
        ada_str = ada_val if (ada_val and ada_val.lower() not in ['o', '0', 'yok', '-']) else "-"
        parsel_str = parsel_val if (parsel_val and parsel_val.lower() not in ['yok', '-']) else "-"
        
        info["ada_parsel"] = f"Ada: {ada_str} / Parsel: {parsel_str}"
        if sahip_val and sahip_val != "-":
            info["yapi_sahibi"] = sahip_val

        return info
    except Exception as e:
        logging.exception("Tutanak okunurken hata: %s", e)
        return info


@st.cache_data(ttl=60)
def veritabani_yukle():
    if not os.path.exists(EXCEL_VT_YOLU):
        return pd.DataFrame(), pd.DataFrame()
    try:
        df_muellif = pd.read_excel(EXCEL_VT_YOLU, sheet_name=0)
        df_muteahhit = pd.read_excel(EXCEL_VT_YOLU, sheet_name=1)
        return df_muellif, df_muteahhit
    except Exception:
        return pd.DataFrame(), pd.DataFrame()


def render():
    st.title("🏗️ Yıkım Planı ve Yasal Evrak Modülü")
    st.markdown("---")

    df_muellif, df_muteahhit = veritabani_yukle()
    if df_muellif.empty or df_muteahhit.empty:
        st.warning(f"⚠️ '{EXCEL_VT_YOLU}' dosyasından veriler okunamadı. Lütfen kontrol edin.")
        return

    st.subheader("📂 1. Adım: Yapı Bilgi Tutanak / Belge Yükleme")
    tutanak_file = st.file_uploader("Yapı Bilgilerini İçeren Excel Dosyasını Yükleyin:", type=["xlsx", "xls"], key="ana_tutanak_dosyasi")
    
    if tutanak_file is not None:
        if "son_okunan_dosya" not in st.session_state or st.session_state.get("son_okunan_dosya") != tutanak_file.name:
            st.session_state["son_okunan_dosya"] = tutanak_file.name
            st.session_state["yapi_bilgileri"] = read_fenni_mesul_details(tutanak_file)
            st.success("✅ Tutanak başarıyla okundu ve hafızaya alındı!")
    else:
        if "yapi_bilgileri" not in st.session_state:
            st.session_state["yapi_bilgileri"] = {"yapi_adresi": "-", "ada_parsel": "Ada: - / Parsel: -", "il_ilce": "-", "idare": "-", "yapi_sahibi": "-"}

    aktif_bilgi = st.session_state["yapi_bilgileri"]
    st.markdown("---")

    # Yapı Sahibi ve Adres Bilgilerinin Düzenlenebileceği Alan
    with st.expander("✏️ Yapı ve Konum Bilgilerini İncele / Düzenle", expanded=True):
        col_hb1, col_hb2 = st.columns(2)
        aktif_bilgi["yapi_adresi"] = col_hb1.text_input("Yapı Adresi:", value=aktif_bilgi.get("yapi_adresi", "-"))
        aktif_bilgi["ada_parsel"] = col_hb2.text_input("Ada / Parsel Bilgisi:", value=aktif_bilgi.get("ada_parsel", "Ada: - / Parsel: -"))
        
        col_hb3, col_hb4 = st.columns(2)
        aktif_bilgi["yapi_sahibi"] = col_hb3.text_input("Yapı Sahibi / İşveren:", value=aktif_bilgi.get("yapi_sahibi", "-"))
        aktif_bilgi["idare"] = col_hb4.text_input("İlgili İdare (Belediye):", value=aktif_bilgi.get("idare", "Belediye Başkanlığı"))

    st.markdown("---")

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

        sozlesme_suresi = st.number_input("Sözleşme Süresi (Gün):", value=90, key="soz_sure")
        ucret = st.text_input("Anlaşma Ücreti (TL):", value="1500 TL + KDV", key="soz_ucret")

        if st.button("🚀 Sözleşmeyi Oluştur ve İndir", type="primary", key="btn_soz"):
            context = {
                "muellif_adi": m_satir.get("Ad_Soyad"), "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                "muellif_tc": m_satir.get("TC_No"), "muellif_tel": m_satir.get("Telefon"),
                "muteahhit_firma": mut_satir.get("Firma_Unvani"), "muteahhit_yetkili": mut_satir.get("Yetkili_Ad_Soyad"),
                "muteahhit_vno": mut_satir.get("Vergi_No_TC"), "muteahhit_adres": mut_satir.get("Adres"),
                "muteahhit_tel": mut_satir.get("Telefon"), "yapi_adresi": aktif_bilgi.get("yapi_adresi"), 
                "ada_parsel": aktif_bilgi.get("ada_parsel"), "yapi_sahibi": aktif_bilgi.get("yapi_sahibi"),
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

    elif alt_islem == "📜 Fenni Mesul Taahhütnamesi":
        st.subheader("📜 Fenni Mesul Taahhütnamesi Hazırlama")
        secilen_fenni = st.selectbox("Fenni Mesul Seçin:", df_muellif["Ad_Soyad"].tolist(), key="fenni_secim")
        f_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_fenni].iloc[0]

        if st.button("🚀 Taahhütnameyi Oluştur", type="primary", key="btn_fenni"):
            context = {
                "fenni_mesul_adi": f_satir.get("Ad_Soyad"),
                "muellif_tc": f_satir.get("TC_No"),
                "muellif_oda_no": f_satir.get("Oda_Sicil_No"),
                "oda_no": f_satir.get("Oda_Sicil_No"),
                "fenni_adres": f_satir.get("Adres"),
                "telefon": f_satir.get("Telefon"),
                "il_ilce": aktif_bilgi.get("il_ilce", "-"),
                "idare": aktif_bilgi.get("idare", "-"),
                "yapi_adresi": aktif_bilgi.get("yapi_adresi", "-"),
                "ada_parsel": aktif_bilgi.get("ada_parsel", "Ada: - / Parsel: -"),
                "yapi_sahibi": aktif_bilgi.get("yapi_sahibi", "-"),
                "tarih": datetime.date.today().strftime("%d.%m.%Y")
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

    elif alt_islem == "📝 Müellif Taahhütnamesi (İdareye Verilecek - Form 2)":
        st.subheader("📝 Müellif Taahhütnamesi (Form 2)")
        secilen_mue = st.selectbox("Müellif Seçin:", df_muellif["Ad_Soyad"].tolist(), key="form2_mue")
        m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_mue].iloc[0]

        if st.button("🚀 Form 2 Oluştur", type="primary", key="btn_form2"):
            context = {
                "muellif_adi": m_satir.get("Ad_Soyad"),
                "oda_no": m_satir.get("Oda_Sicil_No"),
                "muellif_adres": m_satir.get("Adres"),
                "telefon": m_satir.get("Telefon"),
                "il_ilce": aktif_bilgi.get("il_ilce", "-"),
                "idare": aktif_bilgi.get("idare", "Belediye Başkanlığı"),
                "ada_parsel": aktif_bilgi.get("ada_parsel", "Ada: - / Parsel: -"),
                "yapi_adresi": aktif_bilgi.get("yapi_adresi", "-"),
                "yapi_sahibi": aktif_bilgi.get("yapi_sahibi", "-"),
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

    elif alt_islem == "🏗️ Yıkım Planı Raporu":
        st.subheader("🏗️ Yıkım Planı Raporu Oluşturucu")
        col_mue, col_mut = st.columns(2)
        with col_mue:
            secilen_mue = st.selectbox("Proje Müellifi Seçin:", df_muellif["Ad_Soyad"].tolist(), key="yp_mue")
            m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_mue].iloc[0]
        with col_mut:
            secilen_mut = st.selectbox("Müteahhit Firma Seçin:", df_muteahhit["Firma_Unvani"].tolist(), key="yp_mut")
            mut_satir = df_muteahhit[df_muteahhit["Firma_Unvani"] == secilen_mut].iloc[0]

        col3, col4 = st.columns(2)
        yikim_yontemi = col3.selectbox("Yıkım Yöntemi:", ["Mekanik Yıkım (Ekskavatör)", "Kademeli Yıkım", "Elle + Mekanik Yıkım"], key="yp_yontem")
        muhit = col4.selectbox("Saha Konumu:", ["Meskun Mahal", "Sanayi Bölgesi", "Açık / Kırsal"], key="yp_muhit")

        if st.button("🚀 Yıkım Planı Raporunu Oluştur", type="primary", key="btn_yp"):
            context = {
                "muellif_adi": m_satir.get("Ad_Soyad"), "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                "muellif_tc": m_satir.get("TC_No"), "muellif_tel": m_satir.get("Telefon"),
                "muteahhit_firma": mut_satir.get("Firma_Unvani"), "muteahhit_yetkili": mut_satir.get("Yetkili_Ad_Soyad"),
                "muteahhit_vno": mut_satir.get("Vergi_No_TC"), "muteahhit_adres": mut_satir.get("Adres"),
                "muteahhit_tel": mut_satir.get("Telefon"), "yapi_adresi": aktif_bilgi.get("yapi_adresi"), 
                "ada_parsel": aktif_bilgi.get("ada_parsel"), "yapi_sahibi": aktif_bilgi.get("yapi_sahibi"),
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

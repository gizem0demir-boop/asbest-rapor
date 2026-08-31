import datetime
import os
import re
import logging

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
import openpyxl

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


def temizle_sayi_str(val):
    """Excel'den okunan tam sayıların 102.0 yerine 102 olarak görünmesini sağlar."""
    s = str(val).strip()
    if s.endswith(".0"):
        return s[:-2]
    return s


def read_fenni_mesul_details(tutanak_file):
    info = {
        "yapi_adresi": "-",
        "ada_parsel": "-",
        "il_ilce": "-",
        "idare": "-",
        "yapi_sahibi": "-",
        "mahalle": "-",
        "sokak": "-",
        "kapi_no": "-",
        "toplam_bb_sayisi": "",
        "toplam_kat_sayisi": "",
        "toplam_insaat_alani": "",
        "nitelligi": "",
        "yapi_sinifi": "",
        "yapi_grubu": "",
        "bina_yuksekligi": ""
    }
    try:
        if hasattr(tutanak_file, "seek"):
            try:
                tutanak_file.seek(0)
            except Exception:
                pass

        xls = pd.ExcelFile(tutanak_file)
        sheet_to_load = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_to_load, header=None)

        ada_val, parsel_val, adres_val, sahip_val = "", "", "", ""

        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row):
                if pd.isna(val):
                    continue
                val_str = temizle_sayi_str(val)
                val_lower = val_str.lower()

                # Adres yakalama
                if "adres" in val_lower or "firma adresi" in val_lower:
                    if ":" in val_str:
                        parcalar = val_str.split(":", 1)
                        if len(parcalar) > 1 and len(parcalar[1].strip()) > 3:
                            adres_val = parcalar[1].strip()
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        adres_val = temizle_sayi_str(row.iloc[c_idx + 1])

                # Ada yakalama
                if "ada" in val_lower:
                    if ":" in val_str:
                        m = re.search(r'(?:ada)[^0-9]*([0-9\w\-]+)', val_str, re.IGNORECASE)
                        if m:
                            val_bulunan = m.group(1)
                            if val_bulunan.lower() not in ['o', 'yok', '']:
                                ada_val = val_bulunan
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        val_yan = temizle_sayi_str(row.iloc[c_idx + 1])
                        if val_yan.lower() not in ['o', 'yok', '-', '']:
                            ada_val = val_yan

                # Parsel yakalama
                if "parsel" in val_lower:
                    if ":" in val_str:
                        m = re.search(r'(?:parsel)[^0-9]*([0-9\w\-]+)', val_str, re.IGNORECASE)
                        if m:
                            val_bulunan = m.group(1)
                            if val_bulunan.lower() not in ['yok', '']:
                                parsel_val = val_bulunan
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        val_yan = temizle_sayi_str(row.iloc[c_idx + 1])
                        if val_yan.lower() not in ['yok', '-', '']:
                            parsel_val = val_yan

                # Yapı Sahibi / İşveren / Firma Adı yakalama
                if any(k in val_lower for k in ["yapi sahibi", "yapı sahibi", "işveren", "firma adı"]):
                    if ":" in val_str:
                        parcalar = val_str.split(":", 1)
                        if len(parcalar) > 1 and len(parcalar[1].strip()) > 1:
                            sahip_val = parcalar[1].strip()
                    if not sahip_val and c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        sahip_val = temizle_sayi_str(row.iloc[c_idx + 1])

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


def _split_ada_parsel(ada_parsel_str):
    """'Ada: X / Parsel: Y' biçiminden ada dan parseli güvenli şekilde çıkarır."""
    if not ada_parsel_str or not isinstance(ada_parsel_str, str):
        return "-", "-"
    try:
        if "/" in ada_parsel_str:
            parts = ada_parsel_str.split("/")
            ada_part = parts[0].replace("Ada:", "").strip()
            parsel_part = parts[1].replace("Parsel:", "").strip()
            ada = ada_part if ada_part else "-"
            parsel = parsel_part if parsel_part else "-"
            return ada, parsel
        m = re.search(r'([0-9]+)', ada_parsel_str)
        if m:
            return m.group(1), "-"
    except Exception:
        pass
    return "-", "-"


def fill_excel_template(template_path: str, context: dict, output_path: str):
    """Openpyxl ile hücre içindeki '{{key}}'leri context sözlüğündeki değerlerle değiştirir."""
    wb = openpyxl.load_workbook(template_path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if not val or not isinstance(val, str):
                    continue
                new_val = val
                for k, v in context.items():
                    ph = "{{" + k + "}}"
                    if ph in new_val:
                        new_val = new_val.replace(ph, str(v) if v is not None else "")
                if new_val != val:
                    cell.value = new_val
    wb.save(output_path)


def render():
    st.title("️ Yıkım Planı ve Yasal Evrak Modülü")
    st.markdown("---")

    df_muellif, df_muteahhit = veritabani_yukle()
    if df_muellif.empty or df_muteahhit.empty:
        st.warning(f"⚠️ '{EXCEL_VT_YOLU}' dosyasından veriler okunamadı. Lütfen kontrol edin.")
        return

    st.subheader("1. Adım: Yapı Bilgi Tutanak / Belge Yükleme")
    tutanak_file = st.file_uploader("Yapı Bilgilerini İçeren Excel Dosyasını Yükleyin:", type=["xlsx", "xls"], key="ana_tutanak_dosyasi")

    if tutanak_file is not None:
        file_id = getattr(tutanak_file, "file_id", getattr(tutanak_file, "name", None))
        if "son_yuklenen_dosya_id" not in st.session_state or st.session_state.get("son_yuklenen_dosya_id") != file_id:
            st.session_state["son_yuklenen_dosya_id"] = file_id
            st.session_state["yapi_bilgileri"] = read_fenni_mesul_details(tutanak_file)
            st.success("✅ Tutanak başarıyla okundu ve hafızaya alındı!")
    else:
        if "yapi_bilgileri" not in st.session_state:
            st.session_state["yapi_bilgileri"] = {
                "yapi_adresi": "-", "ada_parsel": "Ada: - / Parsel: -", "il_ilce": "-", 
                "idare": "-", "yapi_sahibi": "-", "mahalle": "-", "sokak": "-", "kapi_no": "-",
                "toplam_bb_sayisi": "", "toplam_kat_sayisi": "", "toplam_insaat_alani": "",
                "nitelligi": "", "yapi_sinifi": "", "yapi_grubu": "", "bina_yuksekligi": ""
            }

    aktif_bilgi = st.session_state["yapi_bilgileri"]
    st.markdown("---")

    with st.expander("✏️ Yapı ve Konum Bilgilerini İncele / Düzenle", expanded=True):
        col_hb1, col_hb2 = st.columns(2)
        aktif_bilgi["yapi_adresi"] = col_hb1.text_input("Yapı Adresi:", value=aktif_bilgi.get("yapi_adresi", "-"))
        aktif_bilgi["ada_parsel"] = col_hb2.text_input("Ada / Parsel Bilgisi:", value=aktif_bilgi.get("ada_parsel", "Ada: - / Parsel: -"))

        col_hb3, col_hb4 = st.columns(2)
        aktif_bilgi["yapi_sahibi"] = col_hb3.text_input("Yapı Sahibi / İşveren:", value=aktif_bilgi.get("yapi_sahibi", "-"))
        aktif_bilgi["idare"] = col_hb4.text_input("İlgili İdare (Belediye):", value=aktif_bilgi.get("idare", "Belediye Başkanlığı"))

    st.markdown("---")

    alt_islem = st.selectbox(
        "2. Adım: Oluşturulacak Evrak Türünü Seçin:",
        [
            "-- Seçiniz --",
            "Müellif - Müteahhit Yıkım Sözleşmesi",
            "Fenni Mesul Taahhütnamesi",
            "Müellif Taahhütnamesi (İdareye Verilecek - Form 2)",
            "️ Yıkım Planı Raporu",
        ],
    )
    st.markdown("---")

    if alt_islem == "Müellif - Müteahhit Yıkım Sözleşmesi":
        st.subheader("Müellif ve Müteahhit Yıkım Sözleşmesi")
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

        if st.button("Sözleşmeyi Oluştur ve İndir", type="primary", key="btn_soz"):
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
                    st.download_button("Sözleşmeyi İndir", f, file_name="Yikim_Sozlesmesi.docx", key="dl_soz")
                st.success("✅ Sözleşme başarıyla oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: {sablon_yolu}")

    elif alt_islem == "Fenni Mesul Taahhütnamesi":
        st.subheader("Fenni Mesul Taahhütnamesi Hazırlama")
        secilen_fenni = st.selectbox("Fenni Mesul Seçin:", df_muellif["Ad_Soyad"].tolist(), key="fenni_secim")
        f_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_fenni].iloc[0]

        if st.button("Taahhütnameyi Oluştur", type="primary", key="btn_fenni"):
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
                    st.download_button("Taahhütnameyi İndir", f, file_name="Fenni_Mesul_Taahhutnamesi.docx", key="dl_fenni")
                st.success("✅ Fenni Mesul Taahhütnamesi oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: {sablon_yolu}")

    elif alt_islem == "Müellif Taahhütnamesi (İdareye Verilecek - Form 2)":
        st.subheader("Müellif Taahhütnamesi (Form 2)")
        secilen_mue = st.selectbox("Müellif Seçin:", df_muellif["Ad_Soyad"].tolist(), key="form2_mue")
        m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_mue].iloc[0]

        if st.button("Form 2 Oluştur", type="primary", key="btn_form2"):
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
                    st.download_button("Form 2 İndir", f, file_name="Form2_Muellif_Taahhutnamesi.docx", key="dl_form2")
                st.success("✅ Form 2 Taahhütnamesi oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: {sablon_yolu}")

    elif alt_islem == "️ Yıkım Planı Raporu":
        st.subheader("️ Yıkım Planı Raporu Oluşturucu")
        col_mue, col_mut = st.columns(2)
        with col_mue:
            secilen_mue = st.selectbox("Proje Müellifi / Uzman Seçin:", df_muellif["Ad_Soyad"].tolist(), key="yp_mue")
            m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_mue].iloc[0]
        with col_mut:
            secilen_mut = st.selectbox("Müteahhit Firma Seçin:", df_muteahhit["Firma_Unvani"].tolist(), key="yp_mut")
            mut_satir = df_muteahhit[df_muteahhit["Firma_Unvani"] == secilen_mut].iloc[0]

        col3, col4 = st.columns(2)
        yikim_yontemi = col3.selectbox("Yıkım Yöntemi:", ["Mekanik Yıkım (Ekskavatör)", "Kademeli Yıkım", "Elle + Mekanik Yıkım"], key="yp_yontem")
        muhit = col4.selectbox("Saha Konumu:", ["Meskun Mahal", "Sanayi Bölgesi", "Açık / Kırsal"], key="yp_muhit")

        st.markdown("---")
        st.markdown("⚠️ **Fenni Mesul ve Asbest Denetim Bilgileri**")
        col_as1, col_as2 = st.columns(2)
        asbest_durum_secimi = col_as1.selectbox(
            "Asbest Kontrol Durumu:", 
            ["Asbest Raporu Mevcut ve Negatif", "Asbest İncelemesi Yapıldı / Risk Yok", "Ek Asbest Denetimi Gerekli"], 
            key="yp_asbest_durum"
        )
        fenni_mesul_notu = col_as2.text_input(
            "Fenni Mesul Sorumluluk Notu:", 
            value="Tüm fenni mesuliyet ve asbest kontrol denetimleri üstlenilmiştir.", 
            key="yp_fenni_not"
        )

        st.markdown("---")
        st.markdown("### Yapı Teknik Özellikleri (Tutanaktan gelen verileri kontrol edin / düzenleyin)")
        t_col1, t_col2, t_col3 = st.columns(3)
        aktif_bilgi["toplam_bb_sayisi"] = t_col1.text_input("Toplam B. Bölüm Sayısı:", value=str(aktif_bilgi.get("toplam_bb_sayisi", "")), key="yp_toplam_bb")
        aktif_bilgi["toplam_kat_sayisi"] = t_col2.text_input("Toplam Kat Sayısı:", value=str(aktif_bilgi.get("toplam_kat_sayisi", "")), key="yp_toplam_kat")
        aktif_bilgi["toplam_insaat_alani"] = t_col3.text_input("Toplam İnşaat Alanı (m²):", value=str(aktif_bilgi.get("toplam_insaat_alani", "")), key="yp_toplam_alani")

        t_col4, t_col5, t_col6 = st.columns(3)
        aktif_bilgi["nitelligi"] = t_col4.text_input("Niteliği:", value=str(aktif_bilgi.get("nitelligi", "")), key="yp_nitelik")
        aktif_bilgi["yapi_sinifi"] = t_col5.text_input("Yapı Sınıfı / Grubu:", value=str(aktif_bilgi.get("yapi_sinifi", "")), key="yp_sinif")
        aktif_bilgi["bina_yuksekligi"] = t_col6.text_input("Bina Yüksekliği (m):", value=str(aktif_bilgi.get("bina_yuksekligi", "")), key="yp_yukseklik")

        st.markdown("---")
        st.subheader("️ Yıkım Tekniği, Ekipman ve İş Planı Ayarları")
        secilen_teknik = st.radio(
            "Yıkım Tekniğini Seçiniz:",
            ["Elle", "Y. Erişimli // Kompakt Makinalı", "Kule ve Diğer Yüksek Erişimli Vinç", "Patlayıcılarla", "Kimyasal Madde Kullanarak", "Sıcak / Metal Tozuyla Kesim", "Diğer"],
            horizontal=True,
            key="yp_secilen_teknik"
        )

        diger_yontem_detayi = ""
        if secilen_teknik == "Diğer":
            diger_yontem_detayi = st.text_input("Diğer Yıkım Yöntemini Belirtin:", value="PALETLİ EKSKAVATÖR", key="yp_diger_yontem")

        st.markdown("---")
        st.markdown(" **Yıkımda Görevli Kişi ve Ekipmanlar**")
        col_e1, col_e2, col_e3 = st.columns(3)
        personel_sayisi = col_e1.text_input("Personel Sayısı:", value="4" if secilen_teknik != "Elle" else "6", key="yp_personel")
        makine_aparat_sayisi = col_e2.text_input("Makine-Aparat Sayısı:", value="1", key="yp_makine_sayi")
        operator_sayisi = col_e3.text_input("Operatör Sayısı:", value="1" if secilen_teknik != "Elle" else "0", key="yp_operator_sayi")

        col_e4, col_e5, col_e6 = st.columns(3)
        isci_isaret = col_e4.checkbox("Personel Niteliği: İşçi", value=True, key="yp_chk_isci")
        uzman_isaret = col_e5.checkbox("Personel Niteliği: Uzman", value=True, key="yp_chk_uzman")
        makine_aparat_turu = col_e6.text_area("Makine-Aparat Türü:", value="EKSKAVATÖR KOVASI" if secilen_teknik != "Elle" else "İSKELE\nHAVALI KESME ALETİ\nBALYOZ", key="yp_makine_turu")

        col_e7, col_e8 = st.columns(2)
        operator_belgesi_durumu = col_e7.text_input("Operatör Belgesi / Durumu:", value="VAR" if secilen_teknik != "Elle" else "GEREKMEZ", key="yp_op_belge")
        operator_belge_aciklama = col_e8.text_input("Belge Açıklaması / Detayı:", value="YIKIM PLANININ İÇERİSİNDE MEVCUT" if secilen_teknik != "Elle" else "ELLE YIKIM ŞARTLARINA UYGUNDUR", key="yp_op_belge_aciklama")

        st.markdown("---")
        st.markdown(" **Yıkım İş Planı ve Nizam Durumu**")
        col_n1, col_n2 = st.columns(2)
        nizam_durumu = col_n1.selectbox("Binanın Nizam Durumu:", ["Ayrık Nizam", "Bitişik Nizam"], key="yp_nizam")
        toz_baski_cihazi = col_n2.checkbox("Pulverize Sistemli Toz Bastırma Cihazı Kullanılsın mı?", value=True, key="yp_toz_chk")

        st.markdown("---")
        st.markdown("️ **İnşaat ve Yıkıntı Atıkları Miktarları (Ton)**")
        col_at1, col_at2, col_at3 = st.columns(3)
        try:
            atik_tugla_default = float(aktif_bilgi.get("atik_tugla", 38.0))
        except Exception:
            atik_tugla_default = 38.0
        try:
            atik_metal_default = float(aktif_bilgi.get("atik_metal", 77.0))
        except Exception:
            atik_metal_default = 77.0
        try:
            atik_beton_default = float(aktif_bilgi.get("atik_beton", 990.0))
        except Exception:
            atik_beton_default = 990.0

        atik_tugla = col_at1.number_input("Tuğla (17 01 02) Miktarı (Ton):", min_value=0.0, value=atik_tugla_default, step=1.0, key="yp_at_tugla")
        atik_metal = col_at2.number_input("Karışık Metal (17 04 07) Miktarı (Ton):", min_value=0.0, value=atik_metal_default, step=1.0, key="yp_at_metal")
        atik_beton = col_at3.number_input("Beton (17 01 01) Miktarı (Ton):", min_value=0.0, value=atik_beton_default, step=1.0, key="yp_at_beton")

        st.markdown("---")
        st.markdown(" **Yapı Görselleri (Harita Konumu ve Bina Fotoğrafı)**")
        col_f1, col_f2 = st.columns(2)
        konum_dosya = col_f1.file_uploader("Yapının Konumu (Harita Görseli)", type=["png", "jpg", "jpeg"], key="yp_konum")
        bina_foto_dosya = col_f2.file_uploader("Yapının Fotoğrafı", type=["png", "jpg", "jpeg"], key="yp_bina")

        konum_img_path = None
        bina_img_path = None
        if konum_dosya is not None:
            konum_img_path = "temp_yp_konum.jpg"
            with open(konum_img_path, "wb") as f:
                f.write(konum_dosya.getbuffer())
        if bina_foto_dosya is not None:
            bina_img_path = "temp_yp_bina.jpg"
            with open(bina_img_path, "wb") as f:
                f.write(bina_foto_dosya.getbuffer())

        st.markdown("---")
        st.subheader("Rapor Bilgileri ve Belge Üretimi")
        col_r1, col_r2 = st.columns(2)
        rapor_sayisi = col_r1.text_input("Rapor Sayısı:", value="2026-1276", key="yp_rapor_sayisi")
        bugun_tarihi = datetime.date.today().strftime("%d.%m.%Y")
        col_r2.text_input("Rapor Tarihi:", value=bugun_tarihi, disabled=True, key="yp_rapor_tarih")

        if st.button(" Yıkım Planı Raporunu Oluştur", type="primary", key="yp_btn_uret"):
            if nizam_durumu == "Bitişik Nizam":
                is_p1 = "1. Çatıdan başlayarak yukarıdan aşağı gerçekleşecektir. Bitişik cepheler elle yıkılacaktır."
                sorumluluk_alt = "Yıkım yapılmadan 7 gün önce ilgili idare bilgilendirilecektir. 3 gün önce komşu parseller bilgilendirilecektir."
            else:
                is_p1 = "1. Şantiye şefi tüm alanları kontrol edecek, çevrede canlının olmadığını doğrulayacaktır."
                sorumluluk_alt = "Yıkım yapılmadan 7 gün önce ilgili idare bilgilendirilecektir."

            is_p2 = "2. Yıkım esnasında pulverize toz bastırma cihazı ile sulama yapılacaktır." if toz_baski_cihazi else "2. Yıkım esnasında etrafa toz kalkmaması için sulama yapılacaktır."
            is_p3 = "3. Beton ve çelik enkazlar ekskavatörle temizlenerek parsel içi enkaz sahasına aktarılacaktır."
            is_p4 = f"4. Bina {nizam_durumu.lower()}dir."

            ada_val, parsel_val = _split_ada_parsel(aktif_bilgi.get("ada_parsel", ""))

            atik_listesi = [
                {"atik_no": "1", "atik_kod": "17 01 02", "atik_tanim": "TUĞLA", "atik_miktar": str(int(atik_tugla))},
                {"atik_no": "2", "atik_kod": "17 04 07", "atik_tanim": "KARIŞIK METAL", "atik_miktar": str(int(atik_metal))},
                {"atik_no": "3", "atik_kod": "17 01 01", "atik_tanim": "BETON", "atik_miktar": str(int(atik_beton))}
            ]

            context = {
                "rapor_tarihi": bugun_tarihi,
                "rapor_sayisi": rapor_sayisi,
                "il_ilce": aktif_bilgi.get("il_ilce"),
                "mahalle": aktif_bilgi.get("mahalle", "-"),
                "sokak": aktif_bilgi.get("sokak", "-"),
                "site_adi": aktif_bilgi.get("site_adi", ""),
                "kapi_no": aktif_bilgi.get("kapi_no", "-"),
                " ada": ada_val,
                "parsel": parsel_val,
                "toplam_bb_sayisi": aktif_bilgi.get("toplam_bb_sayisi"),
                "toplam_kat_sayisi": aktif_bilgi.get("toplam_kat_sayisi"),
                "toplam_insaat_alani": aktif_bilgi.get("toplam_insaat_alani"),
                "nitelligi": aktif_bilgi.get("nitelligi"),
                "yapi_sinifi": aktif_bilgi.get("yapi_sinifi"),
                "yapi grubu": aktif_bilgi.get("yapi_sinifi"), # Şablondaki boşluklu anahtar için
                "bina_yuksekligi": aktif_bilgi.get("bina_yuksekligi"),

                # Teknik Seçim İşaretleri (Seçilene "X", diğerlerine boşluk atanır)
                "tekni_elle": "X" if secilen_teknik == "Elle" else "",
                "tekni_kompakt": "X" if secilen_teknik == "Y. Erişimli // Kompakt Makinalı" else "",
                "tekni_kule": "X" if secilen_teknik == "Kule ve Diğer Yüksek Erişimli Vinç" else "",
                "tekni_patlayici": "X" if secilen_teknik == "Patlayıcılarla" else "",
                "tekni_kimyasal": "X" if secilen_teknik == "Kimyasal Madde Kullanarak" else "",
                "tekni_sicak": "X" if secilen_teknik == "Sıcak / Metal Tozuyla Kesim" else "",
                "tekni_diger": "X" if secilen_teknik == "Diğer" else "",
                "yikim_yontemi": diger_yontem_detayi if secilen_teknik == "Diğer" else yikim_yontemi,

                "personel_sayisi": personel_sayisi,
                "makine_aparat_sayisi": makine_aparat_sayisi,
                "operator_sayisi": operator_sayisi,
                "pers_isci": "X" if isci_isaret else " ",
                "pers_uzman": "X" if uzman_isaret else " ",
                "makine_aparat_turu": makine_aparat_turu,
                "operator_belgesi": operator_belgesi_durumu,
                "operator_belge_aciklama": operator_belge_aciklama,

                # Ek makine sayaç boşlukları (şablondaki kalıntılar için)
                "makine_sayisi_1": "", "makine_sayisi_2": "", "makine_sayisi_3": "", "makine_sayisi_4": "",

                "is_plani_1": is_p1,
                "is_plani_2": is_p2,
                "is_plani_3": is_p3,
                "is_plani_4": is_p4,
                "onay_kutusu_1": "X", "onay_kutusu_2": "X", "onay_kutusu_3": "X",
                "sorumluluk_1": "", "sorumluluk_2": "", "sorumluluk_3": "",
                "sorumluluk_alt_aciklama": sorumluluk_alt,

                # Atık Tablosu Satırları
                "atik_no_1": "1", "atik_kod_1": "17 01 02", "atik_tanim_1": "TUĞLA", "atik_miktar_1": str(int(atik_tugla)),
                "atik_no_2": "2", "atik_kod_2": "17 04 07", "atik_tanim_2": "KARIŞIK METAL", "atik_miktar_2": str(int(atik_metal)),
                "atik_no_3": "3", "atik_kod_3": "17 01 01", "atik_tanim_3": "BETON", "atik_miktar_3": str(int(atik_beton)),

                "muellif_ad": m_satir.get("Ad_Soyad"),
                "muellif_oda_no": m_satir.get("Oda_Sicil_No", ""),
                "muteahhit_unvan": mut_satir.get("Firma_Unvani"),
                "muteahhit_tel": mut_satir.get("Telefon", ""),

                "fenni_mesul_adi": m_satir.get("Ad_Soyad"),
                "fenni_mesul_oda_no": m_satir.get("Oda_Sicil_No", ""),
                "asbest_durum_raporu": f"Yapıda asbest kontrolü gerçekleştirilmiş olup durum: {asbest_durum_secimi}.",
                "asbest_kontrol_durumu": asbest_durum_secimi,
                "fenni_mesul_notu": fenni_mesul_notu,

                "yapi_adresi": aktif_bilgi.get("yapi_adresi"),
                "yapi_sahibi": aktif_bilgi.get("yapi_sahibi"),
                "yapinin_konumu": "[Konum Görseli]" if konum_img_path else "",
                "yapinin_fotografi": "[Bina Fotoğrafı]" if bina_img_path else ""
            }

            sablon_docx = os.path.join("templates", "yikim_plani_sablon.docx")
            sablon_xlsx = os.path.join("templates", "yikim_plani_sablon.xlsx")

            if os.path.exists(sablon_docx):
                try:
                    doc = DocxTemplate(sablon_docx)
                    doc.render(context)
                    cikis = "Yikim_Plani_Raporu.docx"
                    doc.save(cikis)
                    with open(cikis, "rb") as f:
                        st.download_button(" Raporu İndir (.docx)", f, file_name=cikis, key="dl_yp")
                    st.success("✅ Yıkım Planı Raporu başarıyla oluşturuldu!")
                except Exception as e:
                    st.error(f"Rapor oluşturulurken hata: {e}")
                    logging.exception("Rapor üretim hatası: %s", e)
            elif os.path.exists(sablon_xlsx):
                try:
                    cikis_xlsx = "Yikim_Plani_Raporu.xlsx"
                    fill_excel_template(sablon_xlsx, context, cikis_xlsx)
                    with open(cikis_xlsx, "rb") as f:
                        st.download_button(" Raporu İndir (.xlsx)", f, file_name=cikis_xlsx, key="dl_yp_xlsx")
                    st.success("✅ Yıkım Planı Raporu (Excel) başarıyla oluşturuldu!")
                except Exception as e:
                    st.error(f"Excel rapor oluşturulurken hata: {e}")
                    logging.exception("Excel rapor üretim hatası: %s", e)
            else:
                st.error(f"❌ Şablon bulunamadı: {sablon_docx} veya {sablon_xlsx}")

if __name__ == "__main__":
    render()

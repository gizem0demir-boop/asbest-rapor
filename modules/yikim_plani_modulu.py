import datetime
import os
import re
import logging

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm

EXCEL_VT_YOLU = "veritabani.xlsx"
TEMPLATE_DIR = "templates"

def klasorleri_kontrol_et():
    if not os.path.exists(TEMPLATE_DIR):
        os.makedirs(TEMPLATE_DIR, exist_ok=True)

klasorleri_kontrol_et()

def sayiyi_yaziya_cevir(tutar_str):
    try:
        rakamlar = re.findall(r"\d+", str(tutar_str))
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
        bosluklu = re.sub(r"(?<!^)(?=[A-Z])", " ", tam_metin)
        return bosluklu + " Türk Lirası"
    except Exception:
        return tutar_str

def read_fenni_mesul_details(tutanak_file):
    info = {
        "yapi_adresi": "-", "mahalle": "-", "sokak": "-", "site_adi": "",
        "kapi_no": "-", "ada": "-", "parsel": "-", "ada_parsel": "Ada: - / Parsel: -",
        "il_ilce": "-", "idare": "-", "yapi_sahibi": "-", "toplam_bb_sayisi": "",
        "toplam_kat_sayisi": "", "toplam_insaat_alani": "", "nitelligi": "",
        "yapi_sinifi": "", "yapi_grubu": "", "bina_yuksekligi": ""
    }
    if not tutanak_file:
        return info
    try:
        if hasattr(tutanak_file, "seek"):
            tutanak_file.seek(0)

        xls = pd.ExcelFile(tutanak_file)
        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)

        mahalle_val, sokak_val, site_val, kapi_val = "", "", "", ""
        ada_val, parsel_val, sahip_val, il_ilce_val = "", "", "", ""
        toplam_bb_val, toplam_kat_val, toplam_alan_val = "", "", ""
        nitelik_val, yapi_sinif_val, yapi_grup_val, bina_yukseklik_val = "", "", "", ""

        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row):
                if pd.isna(val):
                    continue
                val_str = str(val).strip()
                val_lower = val_str.lower()

                # Esnek hücre tarama: Etrafındaki dolu ilk hücreleri yakala
                def sonraki_dolu_hucreyi_bul(r, start_c):
                    for nc in range(start_c + 1, len(r)):
                        if pd.notna(r.iloc[nc]) and str(r.iloc[nc]).strip() not in ["", "-"]:
                            return str(r.iloc[nc]).strip()
                    return ""

                if any(k in val_lower for k in ["il/ilçe", "il / ilçe", "ilçe"]) and not il_ilce_val:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: il_ilce_val = res
                if "mahalle" in val_lower and not mahalle_val:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: mahalle_val = res
                if ("sokak" in val_lower or "cadde" in val_lower) and not sokak_val:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: sokak_val = res
                if "site adı" in val_lower and not site_val:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: site_val = res
                if ("kapı no" in val_lower or "kapi no" in val_lower) and not kapi_val:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: kapi_val = res
                if "ada" in val_lower and ada_val in ["", "-"]:
                    if ":" in val_str:
                        m = re.search(r"(?:ada)[^0-9]*([0-9\w\-]+)", val_str, re.IGNORECASE)
                        if m: ada_val = m.group(1)
                    else:
                        res = sonraki_dolu_hucreyi_bul(row, c_idx)
                        if res: ada_val = res
                if "parsel" in val_lower and parsel_val in ["", "-"]:
                    if ":" in val_str:
                        m = re.search(r"(?:parsel)[^0-9]*([0-9\w\-]+)", val_str, re.IGNORECASE)
                        if m: parsel_val = m.group(1)
                    else:
                        res = sonraki_dolu_hucreyi_bul(row, c_idx)
                        if res: parsel_val = res
                if any(k in val_lower for k in ["yapi sahibi", "işveren", "firma adı"]) and not sahip_val:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: sahip_val = res
                if "toplam b.bölüm" in val_lower or "b. bölüm" in val_lower:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: toplam_bb_val = res
                if "toplam kat sayısı" in val_lower:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: toplam_kat_val = res
                if "toplam inşaat alanı" in val_lower:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: toplam_alan_val = res
                if val_lower == "niteliği" or val_lower == "niteligi":
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: nitelik_val = res
                if "yapı sınıfı" in val_lower:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: yapi_sinif_val = res
                if "bina yüksekliği" in val_lower:
                    res = sonraki_dolu_hucreyi_bul(row, c_idx)
                    if res: bina_yukseklik_val = res

        if il_ilce_val and il_ilce_val != "-":
            info["il_ilce"] = il_ilce_val
            parcalar = il_ilce_val.split("/")
            if len(parcalar) > 1:
                info["idare"] = f"{parcalar[1].strip()} Belediyesi"
        info["mahalle"] = mahalle_val if mahalle_val else "-"
        info["sokak"] = sokak_val if sokak_val else "-"
        info["site_adi"] = site_val if site_val else ""
        info["kapi_no"] = kapi_val if kapi_val else "-"
        info["yapi_adresi"] = " / ".join([p for p in [info["mahalle"], info["sokak"], info["site_adi"], f'No: {info["kapi_no"]}'] if p and p != "-"])

        info["ada"] = ada_val if ada_val else "-"
        info["parsel"] = parsel_val if parsel_val else "-"
        info["ada_parsel"] = f"Ada: {info['ada']} / Parsel: {info['parsel']}"

        if sahip_val:
            info["yapi_sahibi"] = sahip_val

        info["toplam_bb_sayisi"] = toplam_bb_val
        info["toplam_kat_sayisi"] = toplam_kat_val
        info["toplam_insaat_alani"] = toplam_alan_val
        info["nitelligi"] = nitelik_val
        info["yapi_sinifi"] = yapi_sinif_val
        info["bina_yuksekligi"] = bina_yukseklik_val

        return info
    except Exception as e:
        logging.exception("Tutanak okuma hatası: %s", e)
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
    st.title("🏗️ Yıkım Planı ve Yasal Evrak Modülü (Tam Entegre)")
    st.markdown("---")

    df_muellif, df_muteahhit = veritabani_yukle()
    if df_muellif.empty or df_muteahhit.empty:
        st.warning(f"⚠️ '{EXCEL_VT_YOLU}' dosyasından veriler okunamadı. Lütfen geçerli bir Excel veritabanı yerleştirin.")
        return

    st.subheader("📂 1. Adım: Yapı Bilgi Tutanak / Belge Yükleme")
    tutanak_file = st.file_uploader("Yapı Bilgilerini İçeren Excel Dosyasını Yükleyin:", type=["xlsx", "xls"], key="ana_tutanak_dosyasi")

    if tutanak_file is not None:
        file_id = getattr(tutanak_file, "file_id", tutanak_file.name)
        if "son_yuklenen_dosya_id" not in st.session_state or st.session_state.get("son_yuklenen_dosya_id") != file_id:
            st.session_state["son_yuklenen_dosya_id"] = file_id
            st.session_state["yapi_bilgileri"] = read_fenni_mesul_details(tutanak_file)
            st.success("✅ Tutanak başarıyla okundu ve hafızaya alındı!")
    else:
        if "yapi_bilgileri" not in st.session_state:
            st.session_state["yapi_bilgileri"] = read_fenni_mesul_details(None)

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
        st.markdown("🏢 **Yapı Teknik Özellikleri**")
        t_col1, t_col2, t_col3 = st.columns(3)
        aktif_bilgi["toplam_bb_sayisi"] = t_col1.text_input("Toplam B. Bölüm Sayısı:", value=aktif_bilgi.get("toplam_bb_sayisi", ""))
        aktif_bilgi["toplam_kat_sayisi"] = t_col2.text_input("Toplam Kat Sayısı:", value=aktif_bilgi.get("toplam_kat_sayisi", ""))
        aktif_bilgi["toplam_insaat_alani"] = t_col3.text_input("Toplam İnşaat Alanı (m²):", value=aktif_bilgi.get("toplam_insaat_alani", ""))

        t_col4, t_col5, t_col6 = st.columns(3)
        aktif_bilgi["nitelligi"] = t_col4.text_input("Niteliği:", value=aktif_bilgi.get("nitelligi", ""))
        aktif_bilgi["yapi_sinifi"] = t_col5.text_input("Yapı Sınıfı / Grubu:", value=aktif_bilgi.get("yapi_sinifi", ""))
        aktif_bilgi["bina_yuksekligi"] = t_col6.text_input("Bina Yüksekliği (m):", value=aktif_bilgi.get("bina_yuksekligi", ""))

    st.markdown("---")
    st.subheader("🏗️ Yıkım Tekniği, Ekipman ve İş Planı Ayarları")

    secilen_teknik = st.radio(
        "Yıkım Tekniğini Seçiniz:",
        ["Elle", "Y. Erişimli // Kompakt Makinalı", "Kule ve Diğer Yüksek Erişimli Vinç", "Patlayıcılarla", "Kimyasal Madde Kullanarak", "Sıcak / Metal Tozuyla Kesim", "Diğer"],
        horizontal=True,
        key="ana_secilen_teknik"
    )

    diger_yontem_detayi = ""
    if secilen_teknik == "Diğer":
        diger_yontem_detayi = st.text_input("Diğer Yıkım Yöntemini Belirtin:", value="PALETLİ EKSKAVATÖR", key="diger_metin_input")

    st.markdown("---")
    st.markdown("👥 **Yıkımda Görevli Kişi ve Ekipmanlar**")

    if secilen_teknik == "Elle":
        def_pers_sayisi = "6"
        def_makine_turu = "İSKELE\nHAVALI KESME ALETİ\nBALYOZ\nKOREGA BORU"
        def_makine_sayisi = "1"
        def_operator_sayisi = "0"
        def_belge = "GEREKMEZ"
        def_aciklama = "ELLE YIKIM ŞARTLARINA UYGUNDUR"
    else:
        def_pers_sayisi = "4"
        def_makine_turu = "EKSKAVATÖR KOVASI" if secilen_teknik != "Diğer" else diger_yontem_detayi
        def_makine_sayisi = "1"
        def_operator_sayisi = "1"
        def_belge = "VAR"
        def_aciklama = "YIKIM PLANININ İÇERİSİNDE MEVCUT"

    col_e1, col_e2, col_e3 = st.columns(3)
    personel_sayisi = col_e1.text_input("Personel Sayısı:", value=def_pers_sayisi, key="p_sayi")
    makine_aparat_sayisi = col_e2.text_input("Makine-Aparat Sayısı:", value=def_makine_sayisi, key="m_sayi")
    operator_sayisi = col_e3.text_input("Operatör Sayısı:", value=def_operator_sayisi, key="op_sayi")

    col_e4, col_e5, col_e6 = st.columns(3)
    isci_isaret = col_e4.checkbox("Personel Niteliği: İşçi", value=True, key="chk_isci")
    uzman_isaret = col_e5.checkbox("Personel Niteliği: Uzman", value=True, key="chk_uzman")
    makine_aparat_turu = col_e6.text_area("Makine-Aparat Türü:", value=def_makine_turu, key="txt_makine_turu")

    col_e7, col_e8 = st.columns(2)
    operator_belgesi_durumu = col_e7.text_input("Operatör Belgesi / Durumu:", value=def_belge, key="op_belge")
    operator_belge_aciklama = col_e8.text_input("Belge Açıklaması / Detayı:", value=def_aciklama, key="op_belge_aciklama")

    st.markdown("---")
    st.markdown("📋 **Yıkım İş Planı ve Nizam Durumu**")
    col_n1, col_n2 = st.columns(2)
    nizam_durumu = col_n1.selectbox("Binanın Nizam Durumu:", ["Ayrık Nizam", "Bitişik Nizam"], key="nizam_secim")
    toz_baski_cihazi = col_n2.checkbox("Pulverize Sistemli Toz Bastırma Cihazı Kullanılsın mı?", value=True, key="toz_chk")

    st.markdown("---")
    st.markdown("🏗️ **İnşaat ve Yıkıntı Atıkları Miktarları (Ton)**")
    col_at1, col_at2, col_at3 = st.columns(3)
    atik_tugla = col_at1.number_input("Tuğla (17 01 02) Miktarı (Ton):", min_value=0.0, value=38.0, step=1.0, key="at_tugla")
    atik_metal = col_at2.number_input("Karışık Metal (17 04 07) Miktarı (Ton):", min_value=0.0, value=77.0, step=1.0, key="at_metal")
    atik_beton = col_at3.number_input("Beton (17 01 01) Miktarı (Ton):", min_value=0.0, value=990.0, step=1.0, key="at_beton")

    st.markdown("---")
    st.markdown("📷 **Yapı Görselleri (Harita Konumu ve Bina Fotoğrafı)**")
    col_f1, col_f2 = st.columns(2)
    konum_dosya = col_f1.file_uploader("Yapının Konumu (Harita Görseli)", type=["png", "jpg", "jpeg"], key="up_konum")
    bina_foto_dosya = col_f2.file_uploader("Yapının Fotoğrafı", type=["png", "jpg", "jpeg"], key="up_bina")

    konum_img_path = None
    bina_img_path = None
    if konum_dosya is not None:
        konum_img_path = "temp_konum.jpg"
        with open(konum_img_path, "wb") as f:
            f.write(konum_dosya.getbuffer())
    if bina_foto_dosya is not None:
        bina_img_path = "temp_bina.jpg"
        with open(bina_img_path, "wb") as f:
            f.write(bina_foto_dosya.getbuffer())

    st.markdown("---")
    st.subheader("📄 Rapor Bilgileri ve Belge Üretimi")
    col_r1, col_r2 = st.columns(2)
    rapor_sayisi = col_r1.text_input("Rapor Sayısı:", value="2026-1276", key="rp_sayisi_input")
    bugun_tarihi = datetime.date.today().strftime("%d.%m.%Y")
    col_r2.text_input("Rapor Tarihi:", value=bugun_tarihi, disabled=True)

    # 📌 Evrak türü seçimi en başta yer alıyor
    alt_islem = st.selectbox(
        "📌 Oluşturulacak Evrak / Rapor Türünü Seçin:",
        [
            "-- Seçiniz --",
            "🤝 Müellif - Müteahhit Yıkım Sözleşmesi",
            "📜 Fenni Mesul Taahhütnamesi",
            "📝 Müellif Taahhütnamesi (İdareye Verilecek - Form 2)",
            "🏗️ Yıkım Planı Raporu (Tam Kapsamlı)",
        ],
        key="ana_islem_secimi"
    )
    st.markdown("---")

    # 🎯 SEÇİLEN EVRAK TÜRÜNE GÖRE İLGİLİ AYARLAR VE BUTONLAR ALT KISIMDA AÇILIR
    if alt_islem != "-- Seçiniz --":
        st.markdown(f"### ⚙️ {alt_islem} Ayarları ve İşlemleri")
        
        # Sadece ihtiyaç duyulan seçim alanlarını burada gösteriyoruz
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            secilen_muellif_ad = st.selectbox("Müellif Seçiniz:", df_muellif["Ad_Soyad"].tolist(), key="soz_mue_secim")
            m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_muellif_ad].iloc[0]
        with col_m2:
            secilen_mut_firma = st.selectbox("Müteahhit Firma Seçiniz:", df_muteahhit["Firma_Unvani"].tolist(), key="soz_mut_secim")
            mut_satir = df_muteahhit[df_muteahhit["Firma_Unvani"] == secilen_mut_firma].iloc[0]

        if alt_islem == "🤝 Müellif - Müteahhit Yıkım Sözleşmesi":
            sozlesme_suresi = st.number_input("Sözleşme Süresi (Gün):", value=90, key="soz_sure")
            ucret = st.text_input("Anlaşma Ücreti (TL):", value="1500 TL + KDV", key="soz_ucret")

            if st.button("🚀 Sözleşmeyi Oluştur ve İndir", type="primary", key="btn_soz"):
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
                    "yapi_adresi": aktif_bilgi.get("yapi_adresi"),
                    "ada_parsel": aktif_bilgi.get("ada_parsel"),
                    "yapi_sahibi": aktif_bilgi.get("yapi_sahibi"),
                    "sure": sozlesme_suresi,
                    "ucret": ucret,
                    "ucret_yazi": sayiyi_yaziya_cevir(ucret),
                    "tarih": bugun_tarihi
                }
                sablon_yolu = os.path.join(TEMPLATE_DIR, "yikim_sozlesme_sablon.docx")
                if os.path.exists(sablon_yolu):
                    doc = DocxTemplate(sablon_yolu)
                    doc.render(context)
                    cikis = "Yikim_Sozlesmesi.docx"
                    doc.save(cikis)
                    with open(cikis, "rb") as f:
                        st.download_button("📥 Sözleşmeyi İndir", f, file_name="Yikim_Sozlesmesi.docx", key="dl_soz")
                    st.success("✅ Sözleşme başarıyla oluşturuldu!")
                else:
                    st.error(f"❌ Şablon dosyası bulunamadı: '{sablon_yolu}'.")

        elif alt_islem == "📜 Fenni Mesul Taahhütnamesi":
            if st.button("🚀 Taahhütnameyi Oluştur", type="primary", key="btn_fenni"):
                context = {
                    "fenni_mesul_adi": m_satir.get("Ad_Soyad"),
                    "muellif_tc": m_satir.get("TC_No"),
                    "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                    "fenni_adres": m_satir.get("Adres"),
                    "telefon": m_satir.get("Telefon"),
                    "il_ilce": aktif_bilgi.get("il_ilce", "-"),
                    "idare": aktif_bilgi.get("idare", "-"),
                    "yapi_adresi": aktif_bilgi.get("yapi_adresi", "-"),
                    "ada_parsel": aktif_bilgi.get("ada_parsel", "Ada: - / Parsel: -"),
                    "yapi_sahibi": aktif_bilgi.get("yapi_sahibi", "-"),
                    "tarih": bugun_tarihi
                }
                sablon_yolu = os.path.join(TEMPLATE_DIR, "fenni_mesul_taahhutname_sablon.docx")
                if os.path.exists(sablon_yolu):
                    doc = DocxTemplate(sablon_yolu)
                    doc.render(context)
                    cikis = "Fenni_Mesul_Taahhutnamesi.docx"
                    doc.save(cikis)
                    with open(cikis, "rb") as f:
                        st.download_button("📥 Taahhütnameyi İndir", f, file_name="Fenni_Mesul_Taahhutnamesi.docx", key="dl_fenni")
                    st.success("✅ Fenni Mesul Taahhütnamesi oluşturuldu!")
                else:
                    st.error(f"❌ Şablon dosyası bulunamadı: '{sablon_yolu}'.")

        elif alt_islem == "📝 Müellif Taahhütnamesi (İdareye Verilecek - Form 2)":
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
                    "tarih": bugun_tarihi
                }
                sablon_yolu = os.path.join(TEMPLATE_DIR, "form2_taahhutname_sablon.docx")
                if os.path.exists(sablon_yolu):
                    doc = DocxTemplate(sablon_yolu)
                    doc.render(context)
                    cikis = "Form2_Muellif_Taahhutnamesi.docx"
                    doc.save(cikis)
                    with open(cikis, "rb") as f:
                        st.download_button("📥 Form 2 İndir", f, file_name="Form2_Muellif_Taahhutnamesi.docx", key="dl_form2")
                    st.success("✅ Form 2 Taahhütnamesi oluşturuldu!")
                else:
                    st.error(f"❌ Şablon dosyası bulunamadı: '{sablon_yolu}'.")

        elif alt_islem == "🏗️ Yıkım Planı Raporu (Tam Kapsamlı)":
            if st.button("🚀 Yıkım Planı Raporunu Oluştur", type="primary", key="btn_yikim_plani"):
                if nizam_durumu == "Bitişik Nizam":
                    is_p1 = "1. Çatıdan başlayarak yukarıdan aşağı gerçekleşecektir. Bitişik cepheler elle yıkılacaktır."
                    sorumluluk_alt = "Yıkım yapılmadan 7 gün önce ilgili idare bilgilendirilecektir. 3 gün önce komşu parseller bilgilendirilecektir."
                else:
                    is_p1 = "1. Şantiye şefi tüm alanları kontrol edecek, çevrede canlının olmadığını doğrulayacaktır."
                    sorumluluk_alt = "Yıkım yapılmadan 7 gün önce ilgili idare bilgilendirilecektir."

                is_p2 = "2. Yıkım esnasında pulverize toz bastırma cihazı ile sulama yapılacaktır." if toz_baski_cihazi else "2. Yıkım esnasında etrafa toz kalkmaması için sulama yapılacaktır."
                is_p3 = "3. Beton ve çelik enkazlar ekskavatörle temizlenerek parsel içi enkaz sahasına aktarılacaktır."
                is_p4 = f"4. Bina {nizam_durumu.lower()}dir."

                atik_listesi = [
                    {"atik_no": "1", "atik_kod": "17 01 02", "atik_tanim": "TUĞLA", "atik_miktar": str(int(atik_tugla))},
                    {"atik_no": "2", "atik_kod": "17 04 07", "atik_tanim": "KARIŞIK METAL", "atik_miktar": str(int(atik_metal))},
                    {"atik_no": "3", "atik_kod": "17 01 01", "atik_tanim": "BETON", "atik_miktar": str(int(atik_beton))}
                ]

                context = {
                    "rapor_tarihi": bugun_tarihi,
                    "rapor_sayisi": rapor_sayisi,
                    "il_ilce": aktif_bilgi.get("il_ilce"),
                    "mahalle": aktif_bilgi.get("mahalle"),
                    "sokak": aktif_bilgi.get("sokak"),
                    "site_adi": aktif_bilgi.get("site_adi", ""),
                    "kapi_no": aktif_bilgi.get("kapi_no"),
                    "ada": aktif_bilgi.get("ada"),
                    "parsel": aktif_bilgi.get("parsel"),
                    "toplam_bb_sayisi": aktif_bilgi.get("toplam_bb_sayisi"),
                    "toplam_kat_sayisi": aktif_bilgi.get("toplam_kat_sayisi"),
                    "toplam_insaat_alani": aktif_bilgi.get("toplam_insaat_alani"),
                    "nitelligi": aktif_bilgi.get("nitelligi"),
                    "yapi_sinifi": aktif_bilgi.get("yapi_sinifi"),
                    "yapi_grubu": aktif_bilgi.get("yapi_grubu"),
                    "bina_yuksekligi": aktif_bilgi.get("bina_yuksekligi"),

                    "muellif_ad": m_satir.get("Ad_Soyad"),
                    "muellif_oda_no": m_satir.get("Oda_Sicil_No", ""),
                    "muteahhit_unvan": mut_satir.get("Firma_Unvani"),
                    "muteahhit_tel": mut_satir.get("Telefon", ""),

                    "tekni_elle": "X" if secilen_teknik == "Elle" else " ",
                    "tekni_kompakt": "X" if secilen_teknik == "Y. Erişimli // Kompakt Makinalı" else " ",
                    "tekni_kule": "X" if secilen_teknik == "Kule ve Diğer Yüksek Erişimli Vinç" else " ",
                    "tekni_patlayici": "X" if secilen_teknik == "Patlayıcılarla" else " ",
                    "tekni_kimyasal": "X" if secilen_teknik == "Kimyasal Madde Kullanarak" else " ",
                    "tekni_sicak": "X" if secilen_teknik == "Sıcak / Metal Tozuyla Kesim" else " ",
                    "tekni_diger": "X" if secilen_teknik == "Diğer" else " ",
                    "yikim_yontemi": diger_yontem_detayi,

                    "personel_sayisi": personel_sayisi,
                    "makine_aparat_sayisi": makine_aparat_sayisi,
                    "operator_sayisi": operator_sayisi,
                    "pers_isci": "X" if isci_isaret else " ",
                    "pers_uzman": "X" if uzman_isaret else " ",
                    "makine_aparat_turu": makine_aparat_turu,
                    "operator_belgesi": operator_belgesi_durumu,
                    "operator_belge_aciklama": operator_belge_aciklama,

                    "is_plani_1": is_p1,
                    "is_plani_2": is_p2,
                    "is_plani_3": is_p3,
                    "is_plani_4": is_p4,
                    "onay_kutusu_1": "√",
                    "onay_kutusu_2": "√" if toz_baski_cihazi else " ",
                    "onay_kutusu_3": "√",

                    "sorumluluk_1": "Yıkımdan etkileşecek duvar, dayanma yapısı ve komşu binalar kontrol edildi.",
                    "sorumluluk_2": "Yıkılacak binanın etrafı kaldırım işgali olmaksızın 2.50 m. sac ile çevrildi.",
                    "sorumluluk_3": "Yıkım izin belgesi ve sorumlu bilgileri şantiyeye asılacaktır.",
                    "sorumluluk_alt_aciklama": sorumluluk_alt,

                    "atik_listesi": atik_listesi
                }

                sablon_yolu = os.path.join(TEMPLATE_DIR, "yikim_plani_sablon.docx")
                if os.path.exists(sablon_yolu):
                    doc = DocxTemplate(sablon_yolu)

                    if konum_img_path and os.path.exists(konum_img_path):
                        context["yapinin_konumu"] = InlineImage(doc, konum_img_path, width=Cm(6.5), height=Cm(6.5))
                    else:
                        context["yapinin_konumu"] = ""

                    if bina_img_path and os.path.exists(bina_img_path):
                        context["yapinin_fotografi"] = InlineImage(doc, bina_img_path, width=Cm(6.5), height=Cm(6.5))
                    else:
                        context["yapinin_fotografi"] = ""

                    doc.render(context)
                    cikis_dosyasi = "Yikim_Plani_Raporu.docx"
                    doc.save(cikis_dosyasi)

                    with open(cikis_dosyasi, "rb") as f:
                        st.download_button("📥 Raporu İndir (.docx)", f, file_name="Yikim_Plani_Raporu.docx", key="dl_Rapor")
                    st.success("✅ Yıkım Planı Raporu başarıyla oluşturuldu!")
                else:
                    st.error(f"❌ Şablon dosyası bulunamadı: '{sablon_yolu}'. Lütfen templates klasörüne ekleyin.")

if __name__ == "__main__":
    render()

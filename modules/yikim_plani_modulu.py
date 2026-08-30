import datetime
import os
import re
import logging
from io import BytesIO

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm

EXCEL_VT_YOLU = "veritabani.xlsx"
TEMPLATE_DIR = "templates"


def klasorleri_kontrol_et():
    """Gerekli şablon klasörünün varlığını kontrol eder, yoksa oluşturur."""
    if not os.path.exists(TEMPLATE_DIR):
        os.makedirs(TEMPLATE_DIR, exist_ok=True)


klasorleri_kontrol_et()


def sayiyi_yaziya_cevir(tutar_str):
    """Girilen tutar ifadesindeki sayıları Türkçede yasal evrak formatında yazıya çevirir."""
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


from io import BytesIO

def read_fenni_mesul_details(tutanak_file):
    """Daha agresif/robust tutanak okuyucu.
    - BytesIO ile UploadedFile'ı açar
    - Tüm sheet'leri tarar; hücre bazlı regex aramalar yapar
    - DEBUG çıktısı verir: df.head() ve bulunan eşleşmeleri gösterir
    """
    info = {
        "yapi_adresi": "-",
        "mahalle": "-",
        "sokak": "-",
        "site_adi": "",
        "kapi_no": "-",
        "ada": "-",
        "parsel": "-",
        "ada_parsel": "Ada: - / Parsel: -",
        "il_ilce": "-",
        "idare": "-",
        "yapi_sahibi": "-",
        "toplam_bb_sayisi": "",
        "toplam_kat_sayisi": "",
        "toplam_insaat_alani": "",
        "nitelligi": "",
        "yapi_sinifi": "",
        "yapi_grubu": "",
        "bina_yuksekligi": ""
    }

    if not tutanak_file:
        return info

    try:
        # 1) Güvenli açma
        if hasattr(tutanak_file, "getbuffer"):
            buf = BytesIO(tutanak_file.getbuffer())
            xls = pd.ExcelFile(buf)
        else:
            xls = pd.ExcelFile(tutanak_file)

        # DEBUG: sheet isimleri
        sheet_names = xls.sheet_names
        try:
            st.write("DEBUG: sheets:", sheet_names)
        except Exception:
            logging.debug("DEBUG: sheets: %s", sheet_names)

        # helper: normalize cell
        def norm_cell(x):
            if pd.isna(x):
                return ""
            s = str(x).strip()
            if not s:
                return ""
            if s.lower() in ("-", "yok", "none", "nan"):
                return ""
            # float-like '1646.0' -> '1646'
            m = re.match(r"^(-?\d+)\.0+$", s)
            if m:
                return m.group(1)
            return s

        # regex'ler
        ada_regex = re.compile(r'(?:ada[:\s]*)?([0-9]{1,6})', re.IGNORECASE)
        parsel_regex = re.compile(r'(?:parsel[:\s]*)?([0-9]{1,6})', re.IGNORECASE)
        ada_parsel_combo = re.compile(r'ada[:\s]*([0-9]{1,6}).*parsel[:\s]*([0-9]{1,6})', re.IGNORECASE)
        slash_pair = re.compile(r'\b([0-9]{1,6})\s*[\/\-]\s*([0-9]{1,6})\b')  # e.g. "853 / 20" veya "853-20"

        # scans
        found = {"ada": set(), "parsel": set(), "pairs": []}
        # iterate sheets (prefer sheet 0 but scan all)
        for sname in sheet_names:
            df = pd.read_excel(xls, sheet_name=sname, header=None)
            # show top rows for debug
            try:
                st.write(f"DEBUG: sheet={sname} head:", df.head(6))
            except Exception:
                logging.debug("DEBUG: sheet=%s head: %s", sname, df.head(6).to_json())

            # scan each cell
            for r_idx, row in df.iterrows():
                for c_idx, val in enumerate(row):
                    if pd.isna(val):
                        continue
                    cell = str(val).strip()
                    cell_lower = cell.lower()

                    # direct combo: "Ada: 853 / Parsel: 20" or "853 / 20"
                    m_combo = ada_parsel_combo.search(cell)
                    if m_combo:
                        a = m_combo.group(1)
                        p = m_combo.group(2)
                        found["ada"].add(a)
                        found["parsel"].add(p)
                        found["pairs"].append((a, p, sname, r_idx, c_idx))
                        continue

                    # slash pair
                    m_sl = slash_pair.search(cell)
                    if m_sl:
                        a = m_sl.group(1)
                        p = m_sl.group(2)
                        # heuristik: if one part plausibly > 100 => probably ada/parsel order can vary,
                        # we collect both as candidate pair
                        found["pairs"].append((a, p, sname, r_idx, c_idx))
                        # also add to sets (they may be used individually)
                        found["ada"].add(a)
                        found["parsel"].add(p)
                        continue

                    # labels Ada / Parsel in same cell
                    m_ada = re.search(r'ada[:\s]*([0-9]{1,6})', cell, re.IGNORECASE)
                    if m_ada:
                        found["ada"].add(m_ada.group(1))
                    m_par = re.search(r'parsel[:\s]*([0-9]{1,6})', cell, re.IGNORECASE)
                    if m_par:
                        found["parsel"].add(m_par.group(1))

                    # if cell contains the label only (e.g. 'Ada'), check right neighbor
                    if re.search(r'\b(ada|parsel)\b', cell_lower):
                        # try right neighbor
                        try:
                            neighbor = norm_cell(df.iloc[r_idx, c_idx+1]) if c_idx+1 < df.shape[1] else ""
                        except Exception:
                            neighbor = ""
                        if neighbor:
                            if 'ada' in cell_lower:
                                # normalize neighbor numeric if possible
                                m_n = re.search(r'([0-9]{1,6})', neighbor)
                                if m_n:
                                    found["ada"].add(m_n.group(1))
                            if 'parsel' in cell_lower:
                                m_n = re.search(r'([0-9]{1,6})', neighbor)
                                if m_n:
                                    found["parsel"].add(m_n.group(1))

        # Heuristik seçim: prefer explicit pair, else first ada+parsel set members
        chosen_ada = None
        chosen_parsel = None
        if found["pairs"]:
            # prefer pair where ada != '0' and parsel != '0'
            for a,p, sname, r, c in found["pairs"]:
                if a not in ("0","") and p not in ("0",""):
                    chosen_ada = a
                    chosen_parsel = p
                    break
            if not chosen_ada:
                chosen_ada, chosen_parsel = found["pairs"][0][0], found["pairs"][0][1]
        else:
            if found["ada"]:
                # pick the most common or first
                chosen_ada = sorted(found["ada"], key=lambda x: (-len(x), x))[0]
            if found["parsel"]:
                chosen_parsel = sorted(found["parsel"], key=lambda x: (-len(x), x))[0]

        # Assign to info (fall back to '-')
        info["ada"] = chosen_ada if chosen_ada else "-"
        info["parsel"] = chosen_parsel if chosen_parsel else "-"
        info["ada_parsel"] = f"Ada: {info['ada']} / Parsel: {info['parsel']}"

        # DEBUG - show found candidates
        try:
            st.write("DEBUG: found candidates:", found)
            st.write("DEBUG: chosen ada/parsel:", info["ada"], info["parsel"])
        except Exception:
            logging.debug("DEBUG found candidates: %s", found)
            logging.debug("DEBUG chosen ada/parsel: %s %s", info["ada"], info["parsel"])

        return info

    except Exception as e:
        logging.exception("Tutanak okuma hatası (v2): %s", e)
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

    # Tutanak yükleme (okuma + otomatik seçim)
    st.subheader("📂 1. Adım: Yapı Bilgi Tutanak / Belge Yükleme")
    tutanak_file = st.file_uploader("Yapı Bilgilerini İçeren Excel Dosyasını Yükleyin:", type=["xlsx", "xls"], key="ana_tutanak_dosyasi")

    if tutanak_file is not None:
        file_id = getattr(tutanak_file, "file_id", tutanak_file.name)
        if "son_yuklenen_dosya_id" not in st.session_state or st.session_state.get("son_yuklenen_dosya_id") != file_id:
            st.session_state["son_yuklenen_dosya_id"] = file_id
            st.session_state["yapi_bilgileri"] = read_fenni_mesul_details(tutanak_file)
            st.success("✅ Tutanak başarıyla okundu ve hafızaya alındı!")
            # otomatik olarak Yıkım Planı seçilsin (sadece UI selectbox key ile eşleşir)
            st.session_state["ana_islem_secimi"] = "🏗️ Yıkım Planı Raporu (Tam Kapsamlı)"
    else:
        if "yapi_bilgileri" not in st.session_state:
            st.session_state["yapi_bilgileri"] = read_fenni_mesul_details(None)

    aktif_bilgi = st.session_state["yapi_bilgileri"]
    st.markdown("---")

    # Evrak seçimi (bu selectbox session_state'deki ana_islem_secimi'yi kullanır)
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

    # Müellif / Müteahhit seçimleri (her evrak için ortak)
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        secilen_muellif_ad = st.selectbox("Müellif Seçiniz:", df_muellif["Ad_Soyad"].tolist(), key="soz_mue_secim")
        m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_muellif_ad].iloc[0]
    with col_m2:
        secilen_mut_firma = st.selectbox("Müteahhit Firma Seçiniz:", df_muteahhit["Firma_Unvani"].tolist(), key="soz_mut_secim")
        mut_satir = df_muteahhit[df_muteahhit["Firma_Unvani"] == secilen_mut_firma].iloc[0]

    # ---- Diğer evraklar: sözleşme / taahhütler ----
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
                "tarih": datetime.date.today().strftime("%d.%m.%Y")
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
        secilen_fenni = st.selectbox("Fenni Mesul Seçin:", df_muellif["Ad_Soyad"].tolist(), key="fenni_secim")
        f_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_fenni].iloc[0]

        if st.button("🚀 Taahhütnameyi Oluştur", type="primary", key="btn_fenni"):
            context = {
                "fenni_mesul_adi": f_satir.get("Ad_Soyad"),
                "muellif_tc": f_satir.get("TC_No"),
                "muellif_oda_no": f_satir.get("Oda_Sicil_No"),
                "fenni_adres": f_satir.get("Adres"),
                "telefon": f_satir.get("Telefon"),
                "il_ilce": aktif_bilgi.get("il_ilce", "-"),
                "idare": aktif_bilgi.get("idare", "-"),
                "yapi_adresi": aktif_bilgi.get("yapi_adresi", "-"),
                "ada_parsel": aktif_bilgi.get("ada_parsel", "Ada: - / Parsel: -"),
                "yapi_sahibi": aktif_bilgi.get("yapi_sahibi", "-"),
                "tarih": datetime.date.today().strftime("%d.%m.%Y")
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

    # ---- Yıkım Planı Raporu (tam kapsamlı) - tüm yıkım alanları yalnızca burada gösterilir ----
    elif alt_islem == "🏗️ Yıkım Planı Raporu (Tam Kapsamlı)":
        # Yapı teknik bilgileri düzenleme paneli (tutanaktan gelip düzenlenebilsin)
        st.markdown("### 🏢 Yapı Teknik Özellikleri (Tutanaktan gelen verileri kontrol edin / düzenleyin)")
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

        # personel / makina
        st.markdown("---")
        st.markdown("👥 **Yıkımda Görevli Kişi ve Ekipmanlar**")
        col_e1, col_e2, col_e3 = st.columns(3)
        personel_sayisi = col_e1.text_input("Personel Sayısı:", value="4" if secilen_teknik != "Elle" else "6", key="p_sayi")
        makine_aparat_sayisi = col_e2.text_input("Makine-Aparat Sayısı:", value="1", key="m_sayi")
        operator_sayisi = col_e3.text_input("Operatör Sayısı:", value="1" if secilen_teknik != "Elle" else "0", key="op_sayi")

        col_e4, col_e5, col_e6 = st.columns(3)
        isci_isaret = col_e4.checkbox("Personel Niteliği: İşçi", value=True, key="chk_isci")
        uzman_isaret = col_e5.checkbox("Personel Niteliği: Uzman", value=True, key="chk_uzman")
        makine_aparat_turu = col_e6.text_area("Makine-Aparat Türü:", value="EKSKAVATÖR KOVASI" if secilen_teknik != "Elle" else "İSKELE\nHAVALI KESME ALETİ\nBALYOZ", key="txt_makine_turu")

        col_e7, col_e8 = st.columns(2)
        operator_belgesi_durumu = col_e7.text_input("Operatör Belgesi / Durumu:", value="VAR" if secilen_teknik != "Elle" else "GEREKMEZ", key="op_belge")
        operator_belge_aciklama = col_e8.text_input("Belge Açıklaması / Detayı:", value="YIKIM PLANININ İÇERİSİNDE MEVCUT" if secilen_teknik != "Elle" else "ELLE YIKIM ŞARTLARINA UYGUNDUR", key="op_belge_aciklama")

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

        if st.button("🚀 Yıkım Planı Raporunu Excel/Word Şablonuna Aktar", type="primary", key="btn_uret"):
            # örnek iş planı metinleri
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

                # Müellif ve Müteahhit Bilgileri (Veritabanından)
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

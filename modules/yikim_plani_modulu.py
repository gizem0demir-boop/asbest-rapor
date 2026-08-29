import datetime
import io
import os
import re
import sys
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from utils.excel_parser import read_tutanak_details
from utils.pdf_parser import parse_asbestos_pdf_report

EXCEL_VT_YOLU = "veritabani.xlsx"

SUPPORTED_FILE_TYPES = [
    "xlsx",
    "xls",
    "docx",
    "doc",
    "pdf",
    "jpg",
    "jpeg",
    "png",
]


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


def genisletilmis_tutanak_oku(tutanak_file):
    """Yüklenen dosyayı akıllı yedekli mekanizmalarla tarayarak adres ve ada/parsel bilgilerini çıkarır."""
    if tutanak_file is not None:
        try:
            tutanak_file.seek(0)
        except Exception:
            pass
            
        file_name = getattr(tutanak_file, "name", "").lower()
        file_bytes = tutanak_file.read()
        
        adres = ""
        ada = ""
        parsel = ""
        
        # 1. PDF Dosyası ise
        if file_name.endswith(".pdf"):
            temp_path = "temp_yikim_parse.pdf"
            with open(temp_path, "wb") as f:
                f.write(file_bytes)
            try:
                pdf_data = parse_asbestos_pdf_report(temp_path)
                if isinstance(pdf_data, dict):
                    adres = pdf_data.get("adres", "")
                    ada = str(pdf_data.get("ada", ""))
                    parsel = str(pdf_data.get("parsel", ""))
            except Exception:
                pass
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)

        # 2. Excel veya Tablo Dosyası ise
        elif file_name.endswith((".xlsx", ".xls")):
            try:
                # Önce standart parser'ı dene
                res = read_tutanak_details(io.BytesIO(file_bytes))
                info_dict = {}
                if isinstance(res, tuple) and len(res) >= 1 and isinstance(res[0], dict):
                    info_dict = res[0]
                elif isinstance(res, dict):
                    info_dict = res

                for k in ["adres", "yapi_adresi", "Adres", "Yapı Adresi", "MAHALLE"]:
                    if k in info_dict and info_dict[k]:
                        adres = str(info_dict[k])
                        break

                for k in ["ada", "Ada", "ADA"]:
                    if k in info_dict and info_dict[k]:
                        ada = str(info_dict[k])
                        break

                for k in ["parsel", "Parsel", "PARSEL"]:
                    if k in info_dict and info_dict[k]:
                        parsel = str(info_dict[k])
                        break

                # Eğer parser bulamadıysa ham Excel tablosunu satır satır tarayıp Regex ile yakala
                df_excel = pd.read_excel(io.BytesIO(file_bytes), header=None)
                tum_metinler = []
                for col in df_excel.columns:
                    for val in df_excel[col].dropna():
                        val_str = str(val).strip()
                        tum_metinler.append(val_str)
                        val_lower = val_str.lower()
                        
                        if not adres and any(w in val_lower for w in ["adres", "mahalle", "cad", "sokak", "no:"]):
                            if col + 1 < len(df_excel.columns):
                                yan = str(df_excel.iloc[val.name if hasattr(val, 'name') else 0, col + 1]) if hasattr(val, 'name') else ""
                                # Alternatif hücre taraması
                        
                        if not ada and "ada" in val_lower:
                            rakamlar = re.findall(r'\d+', val_str)
                            if rakamlar:
                                ada = rakamlar[0]

                        if not parsel and "parsel" in val_lower:
                            rakamlar = re.findall(r'\d+', val_str)
                            if rakamlar:
                                parsel = rakamlar[0]

                # Ham hücre yan yana arama mantığı
                for col_idx in range(len(df_excel.columns) - 1):
                    for row_idx in range(len(df_excel)):
                        hucre = str(df_excel.iloc[row_idx, col_idx]).lower()
                        yan_hucre = str(df_excel.iloc[row_idx, col_idx + 1])
                        
                        if not adres and ("adres" in hucre or "mahalle" in hucre or "yapı" in hucre):
                            if yan_hucre and yan_hucre != "nan":
                                adres = yan_hucre
                        if not ada and "ada" in hucre:
                            if yan_hucre and yan_hucre != "nan":
                                ada = re.sub(r'\D', '', yan_hucre) or yan_hucre
                        if not parsel and "parsel" in hucre:
                            if yan_hucre and yan_hucre != "nan":
                                parsel = re.sub(r'\D', '', yan_hucre) or yan_hucre

            except Exception as e:
                st.error(f"Dosya işlenirken hata: {e}")

        ada_parsel_str = f"{ada} Ada {parsel} Parsel" if (ada and parsel and ada != "-" and parsel != "-") else (f"{ada} Ada" if ada else (f"{parsel} Parsel" if parsel else ""))
        
        return {
            "yapi_adresi": adres if adres and adres != "nan" else "",
            "ada_parsel": ada_parsel_str,
        }

    return {"yapi_adresi": "", "ada_parsel": ""}


@st.cache_data(ttl=60)
def veritabani_yukle():
    if not os.path.exists(EXCEL_VT_YOLU):
        st.error(
            f"❌ '{EXCEL_VT_YOLU}' bulunamadı. Lütfen repo dizinine Excel dosyasını ekleyin."
        )
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
        st.warning(
            "⚠️ Excel veritabanından veri okunamadı. Lütfen 'veritabani.xlsx' dosyasını kontrol edin."
        )
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
            secilen_muellif_ad = st.selectbox(
                "Müellif Seçiniz:",
                df_muellif["Ad_Soyad"].tolist(),
                key="soz_mue_secim",
            )
            m_satir = df_muellif[
                df_muellif["Ad_Soyad"] == secilen_muellif_ad
            ].iloc[0]

            st.text_input(
                "Oda Sicil No:",
                value=str(m_satir.get("Oda_Sicil_No", "")),
                disabled=True,
                key="soz_mue_oda"
            )
            st.text_input(
                "TC Kimlik No:", value=str(m_satir.get("TC_No", "")), disabled=True, key="soz_mue_tc"
            )
            st.text_input(
                "Müellif Tel:", value=str(m_satir.get("Telefon", "")), disabled=True, key="soz_mue_tel"
            )

        with col2:
            st.markdown("### 🏢 Müteahhit / İşveren")
            secilen_mut_firma = st.selectbox(
                "Müteahhit Firma Seçiniz:",
                df_muteahhit["Firma_Unvani"].tolist(),
                key="soz_mut_secim",
            )
            mut_satir = df_muteahhit[
                df_muteahhit["Firma_Unvani"] == secilen_mut_firma
            ].iloc[0]

            st.text_input(
                "Yetkili Ad Soyad:",
                value=str(mut_satir.get("Yetkili_Ad_Soyad", "")),
                disabled=True,
                key="soz_mut_yetkili"
            )
            st.text_input(
                "Vergi No / TC:",
                value=str(mut_satir.get("Vergi_No_TC", "")),
                disabled=True,
                key="soz_mut_vno"
            )
            st.text_input(
                "Firma Tel:", value=str(mut_satir.get("Telefon", "")), disabled=True, key="soz_mut_tel"
            )

        st.markdown("### 🗺️ Yapı ve Saha Bilgileri")
        tutanak_file = st.file_uploader(
            "📂 Yapı Bilgi Tutanak / Belge Yükleyin (Excel, Word, PDF, Resim):",
            type=SUPPORTED_FILE_TYPES,
            key="soz_tutanak",
        )

        col3, col4 = st.columns(2)
        
        # Dinamik Alan Değerlerini Oturumda Tutsun
        if "soz_adres_val" not in st.session_state:
            st.session_state["soz_adres_val"] = "Kazım Karabekir Mah. 220. Sok. No: 78 Bağcılar, İstanbul"
        if "soz_ada_val" not in st.session_state:
            st.session_state["soz_ada_val"] = "853 Ada 20 Parsel"

        if tutanak_file:
            yapi_data = genisletilmis_tutanak_oku(tutanak_file)
            if yapi_data.get("yapi_adresi"):
                st.session_state["soz_adres_val"] = yapi_data.get("yapi_adresi")
            if yapi_data.get("ada_parsel"):
                st.session_state["soz_ada_val"] = yapi_data.get("ada_parsel")

        yapi_adresi = col3.text_input("Yapı Adresi:", value=st.session_state["soz_adres_val"], key="soz_adres_inp")
        ada_parsel = col4.text_input("Ada / Parsel:", value=st.session_state["soz_ada_val"], key="soz_ada_inp")

        sozlesme_suresi = st.number_input("Sözleşme Süresi (Gün):", value=90, step=15)
        ucret = st.text_input("Anlaşma Ücreti (TL):", value="1500 TL + KDV")
        ucret_yazi = sayiyi_yaziya_cevir(ucret)

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
                "ucret_yazi": ucret_yazi,
                "tarih": datetime.date.today().strftime("%d.%m.%Y"),
            }

            sablon_yolu = "templates/yikim_sozlesme_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis_yolu = "Yikim_Sozlesmesi_Cikti.docx"
                doc.save(cikis_yolu)

                with open(cikis_yolu, "rb") as f:
                    st.download_button(
                        "📥 Hazır Sözleşneyi İndir",
                        f,
                        file_name="Yikim_Sozlesmesi.docx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        ),
                    )
                st.success("✅ Yıkım Sözleşmesi başarıyla oluşturuldu!")
            else:
                st.error(f"❌ Şablon dosyası bulunamadı: '{sablon_yolu}'")

    # 2. FENNİ MESUL TAAHHÜTNAMESİ
    elif alt_islem == "📜 Fenni Mesul Taahhütnamesi":
        st.subheader("📜 Fenni Mesul Taahhütnamesi Hazırlama")

        secilen_fenni = st.selectbox(
            "Fenni Mesul (Mühendis) Seçin:",
            df_muellif["Ad_Soyad"].tolist(),
            key="fenni_secim",
        )
        f_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_fenni].iloc[0]

        st.info(
            f"Seçilen Fenni Mesul: **{f_satir.get('Ad_Soyad')}** | Oda No:"
            f" **{f_satir.get('Oda_Sicil_No')}** | TC: **{f_satir.get('TC_No')}**"
        )

        tutanak_file = st.file_uploader(
            "📂 Tutanak / Belge Yükleyin (Excel, Word, PDF, Resim):",
            type=SUPPORTED_FILE_TYPES,
            key="fenni_tutanak",
        )

        col1, col2 = st.columns(2)
        if "fenni_adres_val" not in st.session_state:
            st.session_state["fenni_adres_val"] = ""
        if "fenni_ada_val" not in st.session_state:
            st.session_state["fenni_ada_val"] = ""

        if tutanak_file:
            yapi_data = genisletilmis_tutanak_oku(tutanak_file)
            if yapi_data.get("yapi_adresi"):
                st.session_state["fenni_adres_val"] = yapi_data.get("yapi_adresi")
            if yapi_data.get("ada_parsel"):
                st.session_state["fenni_ada_val"] = yapi_data.get("ada_parsel")

        yapi_adresi = col1.text_input("Yapı Adresi:", value=st.session_state["fenni_adres_val"], key="fenni_adres")
        ada_parsel = col2.text_input("Ada / Parsel:", value=st.session_state["fenni_ada_val"], key="fenni_ada")

        if st.button("🚀 Fenni Mesul Taahhütnamesi Oluştur", type="primary"):
            context = {
                "fenni_adi": f_satir.get("Ad_Soyad"),
                "fenni_tc": f_satir.get("TC_No"),
                "fenni_oda_no": f_satir.get("Oda_Sicil_No"),
                "yapi_adresi": yapi_adresi,
                "ada_parsel": ada_parsel,
                "tarih": datetime.date.today().strftime("%d.%m.%Y"),
            }
            sablon_yolu = "templates/fenni_mesul_taahhutname_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis = "Fenni_Mesul_Taahhutnamesi.docx"
                doc.save(cikis)
                with open(cikis, "rb") as f:
                    st.download_button(
                        "📥 Taahhütnameyi İndir",
                        f,
                        file_name="Fenni_Mesul_Taahhutnamesi.docx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        ),
                    )
                st.success("✅ Fenni Mesul Taahhütnamesi oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: '{sablon_yolu}'")

    # 3. MÜELLİF TAAHHÜTNAMESİ (FORM 2)
    elif alt_islem == "📝 Müellif Taahhütnamesi (İdareye Verilecek - Form 2)":
        st.subheader("📝 Müellif Taahhütnamesi (Form 2)")

        secilen_mue = st.selectbox(
            "Müellif Seçin:", df_muellif["Ad_Soyad"].tolist(), key="form2_mue"
        )
        m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_mue].iloc[0]

        idare_adi = st.text_input(
            "İlgili İdare / Belediye Adı:",
            value="Kadıköy Belediye Başkanlığı Yapı Kontrol Müdürlüğü'ne",
        )

        tutanak_file = st.file_uploader(
            "📂 Tutanak / Belge Yükleyin:",
            type=SUPPORTED_FILE_TYPES,
            key="form2_tutanak",
        )
        
        col1, col2 = st.columns(2)
        if "form2_adres_val" not in st.session_state:
            st.session_state["form2_adres_val"] = ""
        if "form2_ada_val" not in st.session_state:
            st.session_state["form2_ada_val"] = ""

        if tutanak_file:
            yapi_data = genisletilmis_tutanak_oku(tutanak_file)
            if yapi_data.get("yapi_adresi"):
                st.session_state["form2_adres_val"] = yapi_data.get("yapi_adresi")
            if yapi_data.get("ada_parsel"):
                st.session_state["form2_ada_val"] = yapi_data.get("ada_parsel")

        yapi_adresi = col1.text_input("Yapı Adresi:", value=st.session_state["form2_adres_val"], key="form2_adres")
        ada_parsel = col2.text_input("Ada / Parsel:", value=st.session_state["form2_ada_val"], key="form2_ada")

        if st.button("🚀 Form 2 Taahhütnamesi Oluştur", type="primary"):
            context = {
                "idare_adi": idare_adi,
                "muellif_adi": m_satir.get("Ad_Soyad"),
                "muellif_tc": m_satir.get("TC_No"),
                "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                "yapi_adresi": yapi_adresi,
                "ada_parsel": ada_parsel,
                "tarih": datetime.date.today().strftime("%d.%m.%Y"),
            }
            sablon_yolu = "templates/form2_taahhutname_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis = "Form2_Muellif_Taahhutnamesi.docx"
                doc.save(cikis)
                with open(cikis, "rb") as f:
                    st.download_button(
                        "📥 Form 2 İndir",
                        f,
                        file_name="Form2_Muellif_Taahhutnamesi.docx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        ),
                    )
                st.success("✅ Form 2 Taahhütnamesi oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: '{sablon_yolu}'")

    # 4. YIKIM PLANI RAPORU
    elif alt_islem == "🏗️ Yıkım Planı Raporu":
        st.subheader("🏗️ Yıkım Planı Raporu Oluşturucu")

        secilen_mue = st.selectbox(
            "Proje Müellifi Seçin:", df_muellif["Ad_Soyad"].tolist(), key="yp_mue"
        )
        m_satir = df_muellif[df_muellif["Ad_Soyad"] == secilen_mue].iloc[0]

        secilen_mut = st.selectbox(
            "Müteahhit Firma Seçin:",
            df_muteahhit["Firma_Unvani"].tolist(),
            key="yp_mut",
        )
        mut_satir = df_muteahhit[
            df_muteahhit["Firma_Unvani"] == secilen_mut
        ].iloc[0]

        tutanak_file = st.file_uploader(
            "📂 Tutanak / Belge Yükleyin:",
            type=SUPPORTED_FILE_TYPES,
            key="yp_tutanak",
        )
        
        col1, col2 = st.columns(2)
        if "yp_adres_val" not in st.session_state:
            st.session_state["yp_adres_val"] = ""
        if "yp_ada_val" not in st.session_state:
            st.session_state["yp_ada_val"] = ""

        if tutanak_file:
            yapi_data = genisletilmis_tutanak_oku(tutanak_file)
            if yapi_data.get("yapi_adresi"):
                st.session_state["yp_adres_val"] = yapi_data.get("yapi_adresi")
            if yapi_data.get("ada_parsel"):
                st.session_state["yp_ada_val"] = yapi_data.get("ada_parsel")

        yapi_adresi = col1.text_input("Yapı Adresi:", value=st.session_state["yp_adres_val"], key="yp_adres")
        ada_parsel = col2.text_input("Ada / Parsel:", value=st.session_state["yp_ada_val"], key="yp_ada")

        col3, col4 = st.columns(2)
        yikim_yontemi = col3.selectbox(
            "Yıkım Yöntemi:",
            ["Mekanik Yıkım (Ekskavatör)", "Kademeli Yıkım", "Elle + Mekanik Yıkım"],
        )
        muhit = col4.selectbox(
            "Saha Konumu:", ["Meskun Mahal", "Sanayi Bölgesi", "Açık / Kırsal"]
        )

        if st.button("🚀 Yıkım Planı Raporunu Oluştur", type="primary"):
            context = {
                "muellif_adi": m_satir.get("Ad_Soyad"),
                "muellif_oda_no": m_satir.get("Oda_Sicil_No"),
                "muteahhit_firma": mut_satir.get("Firma_Unvani"),
                "yapi_adresi": yapi_adresi,
                "ada_parsel": ada_parsel,
                "yikim_yontemi": yikim_yontemi,
                "muhit": muhit,
                "tarih": datetime.date.today().strftime("%d.%m.%Y"),
            }
            sablon_yolu = "templates/yikim_plani_sablon.docx"
            if os.path.exists(sablon_yolu):
                doc = DocxTemplate(sablon_yolu)
                doc.render(context)
                cikis = "Yikim_Plani_Raporu.docx"
                doc.save(cikis)
                with open(cikis, "rb") as f:
                    st.download_button(
                        "📥 Yıkım Planı Raporunu İndir",
                        f,
                        file_name="Yikim_Plani_Raporu.docx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        ),
                    )
                st.success("✅ Yıkım Planı Raporu başarıyla oluşturuldu!")
            else:
                st.error(f"❌ Şablon bulunamadı: '{sablon_yolu}'")

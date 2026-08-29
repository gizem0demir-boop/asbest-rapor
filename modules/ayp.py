from collections import OrderedDict
from datetime import datetime
import os
import re

from docx import Document
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st


# Geliştirilmiş ve Esnek Numune Tutanağı Ayrıştırıcı
def parse_asbest_tutanak(file):
    df_raw = pd.read_excel(file, header=None)

    info = {
        "musteri_adi": "",
        "adres": "",
        "pafta": "-",
        "ada": "-",
        "parsel": "-",
        "teklif_no": "",
    }

    # Tüm hücreleri satır satır ve sütun sütun tarayarak anahtar kelimeleri ve değerleri bulalım
    for r_idx in range(len(df_raw)):
        for c_idx in range(len(df_raw.columns)):
            val = df_raw.iloc[r_idx, c_idx]
            if pd.isna(val):
                continue
            text = str(val).strip()

            # Teklif / Talep Numarası
            if any(
                k in text.lower()
                for k in ["teklif no", "talep no", "rapor no", "iş emri"]
            ):
                # Aynı hücrede mi yoksa yan/alt hücrede mi?
                if ":" in text:
                    parts = text.split(":")
                    if len(parts) > 1 and parts[1].strip():
                        info["teklif_no"] = parts[1].strip()
                elif c_idx + 1 < len(df_raw.columns):
                    neighbor = df_raw.iloc[r_idx, c_idx + 1]
                    if pd.notna(neighbor) and str(neighbor).strip() != "nan":
                        info["teklif_no"] = str(neighbor).strip()

            # Firma Adı / Müşteri
            if any(
                k in text.lower()
                for k in ["firma adı", "müşteri", "kurum adı", "unvanı"]
            ):
                if ":" in text:
                    parts = text.split(":")
                    if len(parts) > 1 and parts[1].strip():
                        info["musteri_adi"] = parts[1].strip()
                elif c_idx + 1 < len(df_raw.columns):
                    neighbor = df_raw.iloc[r_idx, c_idx + 1]
                    if pd.notna(neighbor) and str(neighbor).strip() != "nan":
                        info["musteri_adi"] = str(neighbor).strip()

            # Adres
            if any(
                k in text.lower() for k in ["firma adresi", "proje adresi", "adres"]
            ):
                if ":" in text:
                    parts = text.split(":")
                    if len(parts) > 1 and parts[1].strip():
                        info["adres"] = parts[1].strip()
                elif c_idx + 1 < len(df_raw.columns):
                    neighbor = df_raw.iloc[r_idx, c_idx + 1]
                    if pd.notna(neighbor) and str(neighbor).strip() != "nan":
                        info["adres"] = str(neighbor).strip()

            # Pafta, Ada, Parsel
            if "pafta" in text.lower():
                m = re.search(r"pafta\D*([0-9a-z/-]+)", text, re.IGNORECASE)
                if m:
                    info["pafta"] = m.group(1).strip()
                elif c_idx + 1 < len(df_raw.columns):
                    neighbor = df_raw.iloc[r_idx, c_idx + 1]
                    if pd.notna(neighbor):
                        info["pafta"] = str(neighbor).strip()

            if "ada" in text.lower() and "parsel" not in text.lower():
                m = re.search(r"ada\D*([0-9a-z/-]+)", text, re.IGNORECASE)
                if m:
                    info["ada"] = m.group(1).strip()
                elif c_idx + 1 < len(df_raw.columns):
                    neighbor = df_raw.iloc[r_idx, c_idx + 1]
                    if pd.notna(neighbor):
                        info["ada"] = str(neighbor).strip()

            if "parsel" in text.lower():
                m = re.search(r"parsel\D*([0-9a-z/-]+)", text, re.IGNORECASE)
                if m:
                    info["parsel"] = m.group(1).strip()
                elif c_idx + 1 < len(df_raw.columns):
                    neighbor = df_raw.iloc[r_idx, c_idx + 1]
                    if pd.notna(neighbor):
                        info["parsel"] = str(neighbor).strip()

    return info


# ANA AYP MODÜLÜ FONKSİYONU
def render_ayp_module():
    st.title("🏗️ Asbest Yıkım Planı (AYP) Raporu Oluşturucu")

    # Session State Başlangıç Değerleri
    if "ayp_musteri" not in st.session_state:
        st.session_state["ayp_musteri"] = "ABC İnşaat A.Ş."
    if "ayp_adres_Val" not in st.session_state:
        st.session_state["ayp_adres_Val"] = "İstanbul / Türkiye"
    if "ayp_teklif_val" not in st.session_state:
        st.session_state["ayp_teklif_val"] = "AYP.26.1042"
    if "ayp_pafta" not in st.session_state:
        st.session_state["ayp_pafta"] = "-"
    if "ayp_ada" not in st.session_state:
        st.session_state["ayp_ada"] = "-"
    if "ayp_parsel" not in st.session_state:
        st.session_state["ayp_parsel"] = "-"

    # Şablon Seçimi
    st.markdown("### 📑 AYP Rapor Şablonu Seçimi")
    secilen_ayp_sablonu = st.selectbox(
        "Kullanılacak AYP Şablonunu Belirleyin:",
        options=[
            "Esenyurt Şablonu (sablon_ayp_esenyurt.docx)",
            "Sultanbeyli Şablonu (sablon_ayp_sultanbeyli.docx)",
            "Sultangazi Şablonu (sablon_ayp_sultangazi.docx)",
            "Ton Bazlı Şablon (sablon_ayp_ton.docx)",
        ],
        key="selectbox_ayp_sablonu",
    )

    if "Esenyurt" in secilen_ayp_sablonu:
        aktif_sablon = "sablon_ayp_esenyurt.docx"
        sablon_tipi = "esenyurt"
    elif "Sultanbeyli" in secilen_ayp_sablonu:
        aktif_sablon = "sablon_ayp_sultanbeyli.docx"
        sablon_tipi = "sultanbeyli"
    elif "Sultangazi" in secilen_ayp_sablonu:
        aktif_sablon = "sablon_ayp_sultangazi.docx"
        sablon_tipi = "sultangazi"
    else:
        aktif_sablon = "sablon_ayp_ton.docx"
        sablon_tipi = "ton"

    st.markdown("---")
    st.subheader("📂 Numune Tutanağı Yükleme (Veri Otomasyonu)")
    tutanak_file = st.file_uploader(
        "Proje ve adres bilgilerini çekmek için Numune Tutanağı Excel Dosyasını Yükleyin",
        type=["xlsx", "xls"],
        key="ayp_tutanak_uploader",
    )

    if tutanak_file is not None:
        parsed_info = parse_asbest_tutanak(tutanak_file)

        if parsed_info.get("musteri_adi"):
            st.session_state["ayp_musteri"] = parsed_info["musteri_adi"]
        if parsed_info.get("adres"):
            st.session_state["ayp_adres_Val"] = parsed_info["adres"]
        if parsed_info.get("teklif_no"):
            st.session_state["ayp_teklif_val"] = parsed_info["teklif_no"]
        if parsed_info.get("pafta") and parsed_info["pafta"] != "-":
            st.session_state["ayp_pafta"] = parsed_info["pafta"]
        if parsed_info.get("ada") and parsed_info["ada"] != "-":
            st.session_state["ayp_ada"] = parsed_info["ada"]
        if parsed_info.get("parsel") and parsed_info["parsel"] != "-":
            st.session_state["ayp_parsel"] = parsed_info["parsel"]

        st.success(
            "Tutanaktan firma, adres, teklif ve tapu bilgileri başarıyla çekildi!"
        )

    st.markdown("---")
    st.subheader("🔢 Proje ve Rapor Bilgileri")

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        proje_adi = st.text_input(
            "Proje / Bina Adı",
            value="Kentsel Dönüşüm Yıkım Projesi",
            key="ayp_proje_adi",
        )
        mal_sahibi = st.text_input(
            "Mal Sahibi / Müşteri",
            value=st.session_state["ayp_musteri"],
            key="ayp_mal_sahibi",
        )
        adres_input = st.text_area(
            "Bina / Proje Adresi",
            value=st.session_state["ayp_adres_Val"],
            key="ayp_adres",
            height=80,
        )
    with col_p2:
        ayp_rapor_no = st.text_input(
            "AYP Rapor Numarası",
            value=st.session_state["ayp_teklif_val"],
            key="ayp_rapor_no",
        )
        ayp_tarih = st.text_input(
            "Rapor Tarihi",
            value=datetime.now().strftime("%d.%m.%Y"),
            key="ayp_tarih",
        )
        st.markdown("<br>", unsafe_allow_html=True)
        st.info(
            f"📍 **Tapu Bilgileri:** Pafta: {st.session_state['ayp_pafta']} | Ada: {st.session_state['ayp_ada']} | Parsel: {st.session_state['ayp_parsel']}"
        )

    st.markdown("---")
    st.subheader("⚙️ Şablona Özel Hesaplama Parametreleri")

    calc_sonuclar = {}
    if sablon_tipi in ["sultanbeyli", "sultangazi"]:
        st.info(
            f"📌 **{secilen_ayp_sablonu.split(' ')[0]}** şablonu için alan, kat ve cam hesaplama parametreleri:"
        )
        c_col1, c_col2, c_col3 = st.columns(3)
        with c_col1:
            toplam_yapi_alani = st.number_input(
                "Toplam Yapı Alanı (m²)",
                min_value=10.0,
                value=1250.0,
                step=50.0,
                key=f"{sablon_tipi}_alan",
            )
        with c_col2:
            kat_sayisi = st.number_input(
                "Kat Sayısı",
                min_value=1,
                value=5,
                step=1,
                key=f"{sablon_tipi}_kat",
            )
        with c_col3:
            cam_durumu = st.selectbox(
                "Cam Durumu", options=["Var", "Yok"], key=f"{sablon_tipi}_cam"
            )

        carpim_katsayi = 1.25 if cam_durumu == "Var" else 1.00
        hesaplanan_hafriyat = (
            toplam_yapi_alani * kat_sayisi * 0.15 * carpim_katsayi
        )

        calc_sonuclar["toplam_yapi_alani"] = toplam_yapi_alani
        calc_sonuclar["kat_sayisi"] = kat_sayisi
        calc_sonuclar["cam_durumu"] = cam_durumu
        calc_sonuclar["hesaplanan_deger"] = round(hesaplanan_hafriyat, 2)
        st.success(
            f"🧮 Otomatik Hesaplanan Atık/Hacim Değeri: **{calc_sonuclar['hesaplanan_deger']} m³**"
        )

    else:
        st.info(
            f"📂 **{secilen_ayp_sablonu.split(' ')[0]}** şablonu için özel hesaplama Excel dosyası gereklidir."
        )
        ayp_excel_file = st.file_uploader(
            f"Ayp Hesaplama Dosyasını Yükleyin",
            type=["xlsx", "xls"],
            key=f"excel_{sablon_tipi}",
        )

        if sablon_tipi == "ton":
            ton_miktari = st.number_input(
                "Toplam Malzeme Miktarı (Ton cinsinden)",
                min_value=0.1,
                value=45.5,
                step=0.5,
                key="ton_degeri",
            )
            calc_sonuclar["ton_miktari"] = ton_miktari
        else:
            calc_sonuclar["toplam_yapi_alani"] = st.number_input(
                "Yapı Alanı (m²)",
                min_value=10.0,
                value=850.0,
                step=50.0,
                key="esenyurt_alan",
            )

    st.markdown("---")
    if st.button("🚀 AYP Raporunu Oluştur ve İndir", type="primary", key="btn_ayp_olustur"):
        try:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            template_path = os.path.join(base_dir, "templates", aktif_sablon)
            output_path = os.path.join(
                base_dir, f"cikis_ayp_raporu_{ayp_rapor_no}.docx"
            )

            tpl = DocxTemplate(template_path)

            context = {
                "proje_adi": proje_adi,
                "mal_sahibi": mal_sahibi,
                "adres": adres_input,
                "pafta": st.session_state["ayp_pafta"],
                "ada": st.session_state["ayp_ada"],
                "parsel": st.session_state["ayp_parsel"],
                "rapor_no": ayp_rapor_no,
                "rapor_tarihi": ayp_tarih,
                "sablon_turu": sablon_tipi,
            }

            if sablon_tipi in ["sultanbeyli", "sultangazi"]:
                context["toplam_yapi_alani"] = calc_sonuclar["toplam_yapi_alani"]
                context["kat_sayisi"] = calc_sonuclar["kat_sayisi"]
                context["cam_durumu"] = calc_sonuclar["cam_durumu"]
                context["hesaplanan_deger"] = calc_sonuclar["hesaplanan_deger"]
            elif sablon_tipi == "ton":
                context["ton_miktari"] = calc_sonuclar["ton_miktari"]
            elif sablon_tipi == "esenyurt":
                context["toplam_yapi_alani"] = calc_sonuclar.get(
                    "toplam_yapi_alani", 850.0
                )

            tpl.render(context)
            tpl.save(output_path)

            st.success("AYP Raporu başarıyla oluşturuldu!")

            with open(output_path, "rb") as file:
                st.download_button(
                    label="📥 Oluşturulan AYP Raporunu İndir (.docx)",
                    data=file,
                    file_name=f"AYP_Raporu_{ayp_rapor_no}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        except Exception as e:
            st.error(f"AYP Raporu oluşturulurken hata meydana geldi: {e}")

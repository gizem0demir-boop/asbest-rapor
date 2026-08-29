from collections import OrderedDict
from datetime import datetime
import os
import re

from docx import Document
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st


# Tutanağın üst bilgi ve proje bilgilerini okuyan fonksiyon
def parse_asbest_tutanak(file):
    df_raw = pd.read_excel(file, header=None)

    info = {
        "musteri_adi": "ABC İnşaat",
        "adres": "-",
        "pafta": "-",
        "ada": "-",
        "parsel": "-",
        "numune_tarihi": datetime.now().strftime("%d.%m.%Y"),
        "teklif_no": "26-08-5191",
        "telefon": "-",
    }

    for idx in range(min(15, len(df_raw))):
        row_values = [str(x) for x in df_raw.iloc[idx].values if pd.notna(x)]
        row_text = " ".join(row_values)

        if "Talep Numarası" in row_text or "Teklif" in row_text:
            if idx + 1 < len(df_raw):
                for val_candidate in df_raw.iloc[idx + 1].values:
                    if pd.notna(val_candidate) and str(val_candidate).strip() != "nan":
                        info["teklif_no"] = str(val_candidate).strip()
                        break

        if "Firma Adı:" in row_text:
            m = re.search(r"Firma Adı:\s*(.*?)(?:Telefon|$)", row_text)
            if m and m.group(1).strip():
                info["musteri_adi"] = m.group(1).strip()

        if "Firma Adresi:" in row_text:
            m = re.search(r"Firma Adresi:\s*(.*)", row_text)
            if m and m.group(1).strip():
                info["adres"] = m.group(1).strip()

        if "Pafta No:" in row_text or "Parsel No:" in row_text or "Ada No:" in row_text:
            p = re.search(r"Pafta\s*No:\s*([^\s|]*)(?=\s*Ada|$)", row_text, re.IGNORECASE)
            a = re.search(r"Ada\s*No:\s*([^\s|]*)(?=\s*Parsel|$)", row_text, re.IGNORECASE)
            pr = re.search(r"Parsel\s*No:\s*([^\s|]*)(?=$)", row_text, re.IGNORECASE)

            if p and p.group(1).strip():
                info["pafta"] = p.group(1).strip()
            if a and a.group(1).strip():
                info["ada"] = a.group(1).strip()
            if pr and pr.group(1).strip():
                info["parsel"] = pr.group(1).strip()

    return info


# ANA AYP MODÜLÜ FONKSİYONU
def render_ayp_module():
    st.title("🏗️ Asbest Yıkım Planı (AYP) Raporu Oluşturucu")

    # 1. Şablon Seçimi Alanı
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

    # Varsayılan veya tutanaktan okunan veriler
    tutanak_info = {
        "musteri_adi": "ABC İnşaat A.Ş.",
        "adres": "İstanbul / Türkiye",
        "pafta": "-",
        "ada": "-",
        "parsel": "-",
        "teklif_no": "AYP.26.1042",
    }

    if tutanak_file is not None:
        parsed_info = parse_asbest_tutanak(tutanak_file)
        if parsed_info["musteri_adi"] != "ABC İnşaat":
            tutanak_info["musteri_adi"] = parsed_info["musteri_adi"]
        if parsed_info["adres"] != "-":
            tutanak_info["adres"] = parsed_info["adres"]
        tutanak_info["pafta"] = parsed_info["pafta"]
        tutanak_info["ada"] = parsed_info["ada"]
        tutanak_info["parsel"] = parsed_info["parsel"]
        tutanak_info["teklif_no"] = parsed_info["teklif_no"]
        st.success("Tutanaktan proje ve adres bilgileri başarıyla çekildi!")

    st.markdown("---")
    st.subheader("🔢 Proje ve Rapor Bilgileri")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        proje_adi = st.text_input("Proje / Bina Adı", value="Kentsel Dönüşüm Yıkım Projesi", key="ayp_proje_adi")
        mal_sahibi = st.text_input("Mal Sahibi / Müşteri", value=tutanak_info["musteri_adi"], key="ayp_mal_sahibi")
        adres_input = st.text_input("Bina / Proje Adresi", value=tutanak_info["adres"], key="ayp_adres")
    with col_p2:
        ayp_rapor_no = st.text_input("AYP Rapor Numarası", value="AYP.26.1042", key="ayp_rapor_no")
        ayp_tarih = st.text_input("Rapor Tarihi", value=datetime.now().strftime("%d.%m.%Y"), key="ayp_tarih")
        pafta_ada_parsel = f"Pafta: {tutanak_info['pafta']} | Ada: {tutanak_info['ada']} | Parsel: {tutanak_info['parsel']}"
        st.info(f"**Tapu Bilgileri:** {pafta_ada_parsel}")

    st.markdown("---")
    st.subheader("⚙️ Şablona Özel Hesaplama Parametreleri")

    calc_sonuclar = {}
    if sablon_tipi in ["sultanbeyli", "sultangazi"]:
        st.info(f"📌 **{secilen_ayp_sablonu.split(' ')[0]}** şablonu için alan, kat ve cam hesaplama parametreleri:")
        c_col1, c_col2, c_col3 = st.columns(3)
        with c_col1:
            toplam_yapi_alani = st.number_input("Toplam Yapı Alanı (m²)", min_value=10.0, value=1250.0, step=50.0, key=f"{sablon_tipi}_alan")
        with c_col2:
            kat_sayisi = st.number_input("Kat Sayısı", min_value=1, value=5, step=1, key=f"{sablon_tipi}_kat")
        with c_col3:
            cam_durumu = st.selectbox("Cam Durumu", options=["Var", "Yok"], key=f"{sablon_tipi}_cam")

        carpim_katsayi = 1.25 if cam_durumu == "Var" else 1.00
        hesaplanan_hafriyat = toplam_yapi_alani * kat_sayisi * 0.15 * carpim_katsayi
        
        calc_sonuclar["toplam_yapi_alani"] = toplam_yapi_alani
        calc_sonuclar["kat_sayisi"] = kat_sayisi
        calc_sonuclar["cam_durumu"] = cam_durumu
        calc_sonuclar["hesaplanan_deger"] = round(hesaplanan_hafriyat, 2)

        st.success(f"🧮 Otomatik Hesaplanan Atık/Hacim Değeri: **{calc_sonuclar['hesaplanan_deger']} m³** (Cam Katsayısı: {carpim_katsayi})")

    else:
        st.info(f"📂 **{secilen_ayp_sablonu.split(' ')[0]}** şablonu için özel hesaplama Excel dosyası gereklidir.")
        ayp_excel_file = st.file_uploader(
            f"Ayp Hesaplama Dosyasını Yükleyin ({'Ayp Hesaplama Esenyurt' if sablon_tipi=='esenyurt' else 'Ayp Hesaplama Ton'})", 
            type=["xlsx", "xls"],
            key=f"excel_{sablon_tipi}"
        )
        
        if sablon_tipi == "ton":
            ton_miktari = st.number_input("Toplam Malzeme Miktarı (Ton cinsinden)", min_value=0.1, value=45.5, step=0.5, key="ton_degeri")
            calc_sonuclar["ton_miktari"] = ton_miktari
        else:
            calc_sonuclar["ton_miktari"] = 0.0

    st.markdown("---")
    if st.button("🚀 AYP Raporunu Oluştur ve İndir", type="primary", key="btn_ayp_olustur"):
        try:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            template_path = os.path.join(base_dir, "templates", aktif_sablon)
            output_path = os.path.join(base_dir, f"cikis_ayp_raporu_{ayp_rapor_no}.docx")

            tpl = DocxTemplate(template_path)

            context = {
                "proje_adi": proje_adi,
                "mal_sahibi": mal_sahibi,
                "adres": adres_input,
                "pafta": tutanak_info["pafta"],
                "ada": tutanak_info["ada"],
                "parsel": tutanak_info["parsel"],
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

import io
import os
import re
from datetime import datetime
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st
from utils import UPLOAD_FOLDER, read_tutanak_details


def parse_turkish_float(val, default=0.0):
    """Her türlü sayısal veri tipini (float, int, Türkçe string) güvenle float'a dönüştürür."""
    if val is None or pd.isna(val):
        return default
    if isinstance(val, (int, float)):
        return float(val)
    try:
        val_str = str(val).strip()
        # Eğer hem nokta hem virgül varsa (Örn: 1.484,10)
        if "." in val_str and "," in val_str:
            val_str = val_str.replace(".", "").replace(",", ".")
        # Sadece virgül varsa (Örn: 1484,10)
        elif "," in val_str:
            val_str = val_str.replace(",", ".")
        return float(val_str)
    except Exception:
        return default


def format_num(val, decimal_places=2):
    """Sayıları Türkçe formatta (örneğin 1.484,10) stringe dönüştürür."""
    if val is None:
        return "0,00"
    try:
        f_val = parse_turkish_float(val, default=0.0)
        return f"{f_val:,.{decimal_places}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)


def get_float_cell(df, row_idx, col_idx, default=0.0):
    """Excel hücresinden güvenli şekilde float değer okur."""
    try:
        if df.shape[0] > row_idx and df.shape[1] > col_idx:
            val = df.iloc[row_idx, col_idx]
            return parse_turkish_float(val, default=default)
    except Exception:
        pass
    return default


def render_ayp_module():
    st.subheader("♻️ Atık Yönetim Planı (AYP) Rapor Oluşturucu")

    # 1. ŞABLON SEÇİMİ
    st.markdown("### 📑 AYP Rapor Şablonu ve Belediye Seçimi")
    secilen_sablon = st.selectbox(
        "Kullanılacak AYP Şablonunu Belirleyin:",
        options=[
            "Standart AYP Şablonu (sablon_ayp.docx)",
            "Esenyurt AYP Şablonu (sablon_ayp_esenyurt.docx)",
            "Sultanbeyli AYP Şablonu (sablon_ayp_sultanbeyli.docx)",
            "Sultangazi AYP Şablonu (sablon_ayp_sultangazi.docx)",
            "Ton Bazlı AYP Şablonu (sablon_ayp_ton.docx)",
        ],
        key="ayp_sablon_secimi",
    )

    # Seçilen Şablona Göre Bayrakları Belirleme
    is_esenyurt = "Esenyurt" in secilen_sablon
    is_sultanbeyli = "Sultanbeyli" in secilen_sablon
    is_sultangazi = "Sultangazi" in secilen_sablon
    is_ton = "Ton Bazlı" in secilen_sablon

    if is_esenyurt:
        aktif_sablon_dosyasi = "sablon_ayp_esenyurt.docx"
        excel_beklenen_label = "📂 2. AYP Hesaplama Esenyurt Dosyası (Excel):"
    elif is_sultanbeyli:
        aktif_sablon_dosyasi = "sablon_ayp_sultanbeyli.docx"
        excel_beklenen_label = "📂 2. AYP Hesaplama Dosyası (Excel):"
    elif is_sultangazi:
        aktif_sablon_dosyasi = "sablon_ayp_sultangazi.docx"
        excel_beklenen_label = "📂 2. AYP Hesaplama Dosyası (Excel):"
    elif is_ton:
        aktif_sablon_dosyasi = "sablon_ayp_ton.docx"
        excel_beklenen_label = "📂 2. AYP Hesaplama Ton Dosyası (Excel):"
    else:
        aktif_sablon_dosyasi = "sablon_ayp.docx"
        excel_beklenen_label = "📂 2. AYP Hesaplama Dosyası (Excel):"

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        tutanak_file = st.file_uploader(
            "📂 1. Tutanak Dosyası (Excel - Künye için):",
            type=["xlsx", "xls"],
            key="ayp_tutanak",
        )
    with col2:
        ayp_file = st.file_uploader(
            excel_beklenen_label,
            type=["xlsx", "xls"],
            key="ayp_excel",
        )

    # 2. BELEDİYE VE ŞABLONA ÖZEL EK GİRDİLER (Sultanbeyli / Sultangazi / Ton)
    toplam_yapi_alani = 0.0
    kat_sayisi = 6.0
    cam_durumu = "Var"

    if is_sultanbeyli or is_sultangazi or is_ton:
        st.markdown("### 📋 Şablona Özel Parametre ve Girdiler")
        g_col1, g_col2, g_col3 = st.columns(3)
        with g_col1:
            toplam_yapi_alani = st.number_input(
                "Toplam Yapı Alanı (m²)", value=510.0, step=10.0, key="ayp_toplam_yapi_alani"
            )
        with g_col2:
            kat_sayisi = st.number_input(
                "Kat Sayısı", value=6.0, step=1.0, key="ayp_kat_sayisi_ozel"
            )
        with g_col3:
            if is_sultanbeyli or is_sultangazi:
                cam_durumu = st.radio(
                    "Cam Var / Yok Durumu:",
                    options=["Var", "Yok"],
                    horizontal=True,
                    key="ayp_cam_durumu",
                )

    if tutanak_file and ayp_file:
        try:
            os.makedirs(UPLOAD_FOLDER, exist_ok=True)

            # Tutanak dosyasını okuma ve kaydetme
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())

            raw_info = read_tutanak_details(tutanak_path)

            info = {}
            if isinstance(raw_info, tuple):
                if len(raw_info) > 0 and isinstance(raw_info[0], dict):
                    info = raw_info[0].copy()
                if len(raw_info) > 1 and isinstance(raw_info[1], list):
                    info["numuneler"] = raw_info[1]
            elif isinstance(raw_info, dict):
                info = raw_info.copy()

            # AYP hesaplama dosyasını okuma ve kaydetme
            ayp_path = os.path.join(UPLOAD_FOLDER, ayp_file.name)
            with open(ayp_path, "wb") as f:
                f.write(ayp_file.getbuffer())

            # Engine seçimi (xls / xlsx uyumluluğu için)
            excel_engine = "xlrd" if ayp_path.lower().endswith(".xls") else "openpyxl"

            xls = pd.ExcelFile(ayp_path, engine=excel_engine)
            df_sayfa1 = (
                pd.read_excel(ayp_path, sheet_name="Sayfa1", header=None, engine=excel_engine)
                if "Sayfa1" in xls.sheet_names
                else pd.DataFrame()
            )
            df_sayfa2 = (
                pd.read_excel(ayp_path, sheet_name="Sayfa2", header=None, engine=excel_engine)
                if "Sayfa2" in xls.sheet_names
                else pd.DataFrame()
            )

            # --- Sayfa1 Dinamik Hücre Okumaları ---
            # Taban / Kat Alanı (Sayfa1 Kat Alanı G16 hücresi - index row 15, col 6)
            taban_alani = get_float_cell(df_sayfa1, 15, 6, default=85.0)

            # Çatı Alanı ve Kiremit Ağırlığı (Sayfa1 Row 27 - G27 Çatı Alanı, H27 Kiremit Ağırlığı)
            cati_alani = get_float_cell(df_sayfa1, 27, 6, default=0.0)
            kiremit_kg = get_float_cell(df_sayfa1, 27, 7, default=0.0)

            # Seramik Miktarı (Sayfa1 Row 32, I32 / J32)
            seramik_toplam_degeri = get_float_cell(df_sayfa1, 32, 8, default=2684.1)

            # Ahşap Miktarı (Sayfa1 Row 25, J25)
            ahsap_kg_excel = get_float_cell(df_sayfa1, 25, 9, default=792.0)

            # Tuğla Miktarı (Sayfa1 Row 6, J6)
            tugla_kg_excel = get_float_cell(df_sayfa1, 6, 9, default=21780.0)

            # Sıva Miktarı (Sayfa1 Row 10, I10)
            siva_kg_excel = get_float_cell(df_sayfa1, 10, 8, default=72600.0)

            # Karışım Miktarı (SADECE ESENYURT ŞABLONU İÇİN - Sayfa2 H15 Hücresi)
            karisim_toplam_kg = 0.0
            if is_esenyurt:
                karisim_toplam_kg = get_float_cell(df_sayfa2, 14, 7, default=473744.10)

            # Sayfa2 Atık Miktarları Parse Mantığı
            atik_miktarlari = {}
            genel_toplam = 0.0

            for idx, row in df_sayfa2.iterrows():
                row_vals = [v for v in row.values if pd.notna(v)]
                if not row_vals:
                    continue

                row_str_full = " ".join([str(v) for v in row_vals]).lower()
                if "tutar" in row_str_full or ("miktar" in row_str_full and idx < 5):
                    continue

                if "toplam" in row_str_full and "daire" not in row_str_full:
                    for v in row.values:
                        val_f = parse_turkish_float(v, default=0.0)
                        if val_f > 0:
                            genel_toplam = val_f
                            break

                key = row.iloc[5] if len(row) > 5 else None
                val = row.iloc[6] if len(row) > 6 else None

                if pd.notna(key) and str(key).strip().lower() != "atık kodu tanımı":
                    val_num = parse_turkish_float(val, default=0.0)
                    atik_miktarlari[str(key).strip().lower()] = val_num

            # Cam Hesabı Mantığı (Sultanbeyli / Sultangazi için Var / Yok seçimi)
            cam_miktari_kg = atik_miktarlari.get("cam ambalaj", 0.0)
            if (is_sultanbeyli or is_sultangazi) and cam_durumu == "Yok":
                cam_miktari_kg = 0.0
                cam_durumu_metni = "Yapıda cam atık bulunmamaktadır."
            else:
                cam_durumu_metni = "Var" if cam_miktari_kg > 0 else "Yok"

            # Sultanbeyli/Sultangazi haricinde taban alanı düzenlemesi
            if is_sultanbeyli or is_sultangazi:
                taban_alani = (
                    toplam_yapi_alani / kat_sayisi if kat_sayisi > 0 else taban_alani
                )

            bugun_tarihi = datetime.now().strftime("%d.%m.%Y")

            asbest_kg = atik_miktarlari.get("asbest içeren inşaat malzemeleri", 0.0)
            beton_kg = atik_miktarlari.get("beton", 449280.0)
            ahsap_kg = atik_miktarlari.get("ahşap", ahsap_kg_excel)
            tugla_kg = atik_miktarlari.get("tuğla", tugla_kg_excel)
            siva_kg = atik_miktarlari.get("17 08 01 dışındaki alçı bazlı inşaat malzemeleri", siva_kg_excel)
            metal_kg = atik_miktarlari.get("karışık metaller", 33280.0)
            kagit_kg = atik_miktarlari.get("kağıt ve karton ambalaj", 12.0)
            plastik_kg = atik_miktarlari.get("plastik ambalaj", 0.0)

            # Genel toplam Excel'den okunamazsa atıkların toplamını hesapla
            hesaplanan_toplam = sum([
                asbest_kg, beton_kg, ahsap_kg, tugla_kg, siva_kg,
                metal_kg, kagit_kg, plastik_kg, cam_miktari_kg, kiremit_kg, seramik_toplam_degeri
            ])
            genel_toplam_kg = genel_toplam if genel_toplam > 0 else hesaplanan_toplam

            # Context Sözlüğü Hazırlığı
            info.update(
                {
                    "tarih": bugun_tarihi,
                    "TARIH": bugun_tarihi,
                    "rapor_tarihi": bugun_tarihi,
                    "toplam_yapi_alani": toplam_yapi_alani,
                    "toplam_yapi_alani_fmt": format_num(toplam_yapi_alani),
                    "alan_m2": taban_alani,
                    "alan_m2_fmt": format_num(taban_alani),
                    "cati_alan_m2": cati_alani,
                    "cati_alan_m2_fmt": format_num(cati_alani),
                    "kat_sayisi": kat_sayisi,
                    "cam_durumu": cam_durumu_metni,
                    # Kg Bazlı Ham Değerler
                    "asbest_toplam_kg": asbest_kg,
                    "beton_toplam_kg": beton_kg,
                    "kiremit_toplam_kg": kiremit_kg,
                    "seramik_genel_toplam_kg": seramik_toplam_degeri,
                    "ahsap_toplam_kg": ahsap_kg,
                    "tugla_toplam_kg": tugla_kg,
                    "siva_toplam_kg": siva_kg,
                    "toplam_karisik_metal": metal_kg,
                    "kagit_toplam_kg": kagit_kg,
                    "plastik_toplam_kg": plastik_kg,
                    "cam_miktari": cam_miktari_kg,
                    "genel_toplam_miktar": genel_toplam_kg,
                    # Kg Formatlı Değerler
                    "asbest_toplam_kg_fmt": format_num(asbest_kg),
                    "beton_toplam_kg_fmt": format_num(beton_kg),
                    "kiremit_toplam_kg_fmt": format_num(kiremit_kg),
                    "seramik_genel_toplam_kg_fmt": format_num(seramik_toplam_degeri),
                    "ahsap_toplam_kg_fmt": format_num(ahsap_kg),
                    "tugla_toplam_kg_fmt": format_num(tugla_kg),
                    "siva_toplam_kg_fmt": format_num(siva_kg),
                    "toplam_karisik_metal_fmt": format_num(metal_kg),
                    "cam_miktari_fmt": format_num(cam_miktari_kg),
                    "genel_toplam_miktar_fmt": format_num(genel_toplam_kg),
                    # TON BAZLI HESAPLAMALAR VE FORMATLAR
                    "asbest_toplam_ton": asbest_kg / 1000.0,
                    "beton_toplam_ton": beton_kg / 1000.0,
                    "kiremit_toplam_ton": kiremit_kg / 1000.0,
                    "seramik_toplam_ton": seramik_toplam_degeri / 1000.0,
                    "ahsap_toplam_ton": ahsap_kg / 1000.0,
                    "tugla_toplam_ton": tugla_kg / 1000.0,
                    "siva_toplam_ton": siva_kg / 1000.0,
                    "metal_toplam_ton": metal_kg / 1000.0,
                    "kagit_toplam_ton": kagit_kg / 1000.0,
                    "plastik_toplam_ton": plastik_kg / 1000.0,
                    "cam_toplam_ton": cam_miktari_kg / 1000.0,
                    "genel_toplam_ton": genel_toplam_kg / 1000.0,
                    "asbest_toplam_ton_fmt": format_num(asbest_kg / 1000.0),
                    "beton_toplam_ton_fmt": format_num(beton_kg / 1000.0),
                    "kiremit_toplam_ton_fmt": format_num(kiremit_kg / 1000.0),
                    "seramik_toplam_ton_fmt": format_num(seramik_toplam_degeri / 1000.0),
                    "ahsap_toplam_ton_fmt": format_num(ahsap_kg / 1000.0),
                    "tugla_toplam_ton_fmt": format_num(tugla_kg / 1000.0),
                    "siva_toplam_ton_fmt": format_num(siva_kg / 1000.0),
                    "metal_toplam_ton_fmt": format_num(metal_kg / 1000.0),
                    "cam_toplam_ton_fmt": format_num(cam_miktari_kg / 1000.0),
                    "genel_toplam_ton_fmt": format_num(genel_toplam_kg / 1000.0),
                }
            )

            # SADECE ESENYURT İÇİN KARIŞIM DEĞİŞKENLERİNİ CONTEXT'E EKLE
            if is_esenyurt:
                info.update(
                    {
                        "karisim_toplam_kg": karisim_toplam_kg,
                        "karisim_toplam_kg_fmt": format_num(karisim_toplam_kg),
                        "karisim_toplam_ton": karisim_toplam_kg / 1000.0,
                        "karisim_toplam_ton_fmt": format_num(karisim_toplam_kg / 1000.0),
                    }
                )

            st.success(f"✅ Tutanak ve Excel verileri ({secilen_sablon}) için hazırlandı.")

            st.markdown("---")
            if st.button("🚀 AYP Raporunu Oluştur ve Hazırla", type="primary", key="btn_ayp_olustur"):
                # Şablon bulma esnek arama
                current_script_dir = os.path.dirname(os.path.abspath(__file__))
                possible_paths = [
                    os.path.join(current_script_dir, "templates", aktif_sablon_dosyasi),
                    os.path.join(current_script_dir, aktif_sablon_dosyasi),
                    os.path.join(os.getcwd(), "templates", aktif_sablon_dosyasi),
                    os.path.join(os.getcwd(), aktif_sablon_dosyasi),
                ]

                template_path = None
                for path in possible_paths:
                    if os.path.exists(path):
                        template_path = path
                        break

                if template_path:
                    doc = DocxTemplate(template_path)
                    doc.render(info)

                    musteri_adi = info.get("musteri_adi", "Musteri")
                    # Dosya adı için güvenli karaktere çevirme
                    safe_musteri_adi = re.sub(r'[\\/*?:"<>|]', "_", str(musteri_adi))
                    output_filename = f"AYP_Raporu_{safe_musteri_adi}.docx"
                    output_path = os.path.join(UPLOAD_FOLDER, output_filename)
                    doc.save(output_path)

                    st.success(f"✅ Atık Yönetim Planı Raporu ({aktif_sablon_dosyasi}) kullanılarak üretildi!")

                    with open(output_path, "rb") as f:
                        file_data = f.read()

                    st.download_button(
                        label="📥 AYP Raporunu İndir (.docx)",
                        data=file_data,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="btn_ayp_download",
                    )
                else:
                    st.error(
                        f"❌ Şablon dosyası bulunamadı: '{aktif_sablon_dosyasi}'. Lütfen dosyanın 'templates' dizininde olduğundan emin olun."
                    )

        except Exception as e:
            st.error(f"❌ AYP raporu işlenirken hata oluştu: {e}")
    else:
        st.info(
            "ℹ️ Lütfen raporu oluşturmak için hem **Tutanak Dosyasını** hem de ilgili **AYP Hesaplama Dosyasını** yükleyin."
        )

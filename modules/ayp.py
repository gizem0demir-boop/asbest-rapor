import io
import os
import re
import zipfile
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from PIL import Image, ImageOps
import pandas as pd
import streamlit as st
from utils import UPLOAD_FOLDER, read_tutanak_details

# --- ŞABLON YAPILANDIRMASI (CONFIG) ---
SABLON_AYARLARI = {
    "Standart AYP Şablonu (sablon_ayp.docx)": {
        "file_name": "sablon_ayp.docx",
        "label": "📂 2. AYP Hesaplama Dosyası (Excel):",
        "has_esenyurt_karisim": False,
        "is_ton_bazli_excel": False,
        "is_sultangazi": False,
        "is_sultanbeyli": False,
        "is_ankara": False,
        "is_pendik": False,
        "requires_excel": True,
    },
    "Ankara AYP Şablonu (sablon_ayp_ankara.docx)": {
        "file_name": "sablon_ayp_ankara.docx",
        "label": "📂 2. Ankara AYP Hesaplama Dosyası (Excel):",
        "has_esenyurt_karisim": False,
        "is_ton_bazli_excel": False,
        "is_sultangazi": False,
        "is_sultanbeyli": False,
        "is_ankara": True,
        "is_pendik": False,
        "requires_excel": True,
    },
    "Esenyurt AYP Şablonu (sablon_ayp_esenyurt.docx)": {
        "file_name": "sablon_ayp_esenyurt.docx",
        "label": "📂 2. AYP Hesaplama Esenyurt Dosyası (Excel):",
        "has_esenyurt_karisim": True,
        "is_ton_bazli_excel": False,
        "is_sultangazi": False,
        "is_sultanbeyli": False,
        "is_ankara": False,
        "is_pendik": False,
        "requires_excel": True,
    },
    "Sultanbeyli AYP Şablonu (sablon_ayp_sultanbeyli.docx)": {
        "file_name": "sablon_ayp_sultanbeyli.docx",
        "label": "📂 2. AYP Hesaplama Sultanbeyli Dosyası (Excel):",
        "has_esenyurt_karisim": False,
        "is_ton_bazli_excel": False,
        "is_sultangazi": False,
        "is_sultanbeyli": True,
        "is_ankara": False,
        "is_pendik": False,
        "requires_excel": False,
    },
    "Sultangazi AYP Şablonu (sablon_ayp_sultangazi.docx)": {
        "file_name": "sablon_ayp_sultangazi.docx",
        "label": "📂 2. AYP Hesaplama Sultangazi Dosyası (Excel):",
        "has_esenyurt_karisim": False,
        "is_ton_bazli_excel": False,
        "is_sultangazi": True,
        "is_sultanbeyli": False,
        "is_ankara": False,
        "is_pendik": False,
        "requires_excel": False,
    },
    "Pendik AYP Şablonu - 1 (sablon_ayp_pendik_1.docx)": {
        "file_name": "sablon_ayp_pendik_1.docx",
        "label": "📂 2. Pendik Hesaplama Dosyası (Excel):",
        "has_esenyurt_karisim": False,
        "is_ton_bazli_excel": False,
        "is_sultangazi": False,
        "is_sultanbeyli": False,
        "is_ankara": False,
        "is_pendik": True,
        "pendik_tip": 1,
        "requires_excel": True,
    },
    "Pendik AYP Şablonu - 2 (sablon_ayp_pendik_2.docx)": {
        "file_name": "sablon_ayp_pendik_2.docx",
        "label": "📂 2. Pendik Hesaplama Dosyası (Excel):",
        "has_esenyurt_karisim": False,
        "is_ton_bazli_excel": False,
        "is_sultangazi": False,
        "is_sultanbeyli": False,
        "is_ankara": False,
        "is_pendik": True,
        "pendik_tip": 2,
        "requires_excel": True,
    },
    "Ton Bazlı AYP Şablonu (sablon_ayp_ton.docx)": {
        "file_name": "sablon_ayp_ton.docx",
        "label": "📂 2. AYP Hesaplama Ton Dosyası (Excel):",
        "has_esenyurt_karisim": False,
        "is_ton_bazli_excel": True,
        "is_sultangazi": False,
        "is_sultanbeyli": False,
        "is_ankara": False,
        "is_pendik": False,
        "requires_excel": True,
    },
}


def muhendisleri_excelden_oku():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    muhendis_path = os.path.join(base_dir, "templates", "muhendisler.xlsx")

    muhendis_dict = {"Seçiniz...": {"sicil": "", "tc": ""}}

    if os.path.exists(muhendis_path):
        try:
            df = pd.read_excel(muhendis_path)
            for _, row in df.iterrows():
                ad = str(row.iloc[0]).strip()
                sicil = str(row.iloc[1]).strip()
                tc = str(row.iloc[2]).strip()
                if ad and ad != "nan":
                    muhendis_dict[ad] = {
                        "sicil": "" if sicil == "nan" else sicil,
                        "tc": "" if tc == "nan" else tc,
                    }
        except Exception as e:
            st.warning(f"⚠️ Mühendis listesi okunurken hata oluştu: {e}")
    else:
        muhendis_dict = {
            "Seçiniz...": {"sicil": "", "tc": ""},
            "Örnek Mühendis (Excel Bulunamadı)": {
                "sicil": "00000",
                "tc": "00000000000",
            },
        }

    return muhendis_dict


def parse_turkish_float(val, default=0.0):
    if val is None or pd.isna(val):
        return float(default)
    if isinstance(val, (int, float)):
        return float(val)
    try:
        val_str = str(val).strip()
        if "." in val_str and "," in val_str:
            val_str = val_str.replace(".", "").replace(",", ".")
        elif "," in val_str:
            val_str = val_str.replace(",", ".")
        return float(val_str)
    except Exception:
        return float(default)


def format_num(val, decimal_places=1):
    if val is None:
        return "0,0"
    try:
        f_val = parse_turkish_float(val, default=0.0)
        return (
            f"{f_val:,.{decimal_places}f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
    except Exception:
        return str(val)


def get_float_cell(df, row_idx, col_idx, default=0.0):
    try:
        r_idx = int(row_idx)
        c_idx = int(col_idx)
        if (
            isinstance(df, pd.DataFrame)
            and df.shape[0] > r_idx
            and df.shape[1] > c_idx
        ):
            val = df.iloc[r_idx, c_idx]
            return parse_turkish_float(val, default=default)
    except Exception:
        pass
    return float(default)


def sanitize_context_for_jinja(context_dict):
    clean_dict = {}
    for key, val in context_dict.items():
        if isinstance(val, str):
            val_stripped = val.strip()
            if val_stripped.isdigit() or (
                val_stripped.startswith("-") and val_stripped[1:].isdigit()
            ):
                clean_dict[key] = int(val_stripped)
            else:
                try:
                    parsed_f = parse_turkish_float(val_stripped, default=None)
                    if parsed_f is not None and (
                        "." in val_stripped or "," in val_stripped
                    ):
                        clean_dict[key] = parsed_f
                    else:
                        clean_dict[key] = val
                except Exception:
                    clean_dict[key] = val
        else:
            clean_dict[key] = val
    return clean_dict


def process_ayp_excel(e_path, has_esenyurt_karisim=False):
    excel_engine = "xlrd" if e_path.lower().endswith(".xls") else "openpyxl"
    xls = pd.ExcelFile(e_path, engine=excel_engine)

    df_sayfa1 = (
        pd.read_excel(e_path, sheet_name="Sayfa1", header=None, engine=excel_engine)
        if "Sayfa1" in xls.sheet_names
        else pd.DataFrame()
    )
    df_sayfa2 = (
        pd.read_excel(e_path, sheet_name="Sayfa2", header=None, engine=excel_engine)
        if "Sayfa2" in xls.sheet_names
        else pd.DataFrame()
    )

    # --- Sayfa 1 Dinamik Hücre Okumaları ---
    # Kuru Beton: Row 16 (idx 15) -> G16=Col6 (m2), I16=Col8 (kat sayısı), J16=Col9 (toplam beton kg)
    alan_m2 = get_float_cell(df_sayfa1, 15, 6, default=208.0)
    kat_sayisi = get_float_cell(df_sayfa1, 15, 8, default=4.0)
    beton_toplg = get_float_cell(df_sayfa1, 15, 9, default=299520.0)

    # Ahşap Laminant Parke: Row 25 (idx 24) -> F25=Col5, G25=Col6, H25=Col7, I25=Col8, J25=Col9
    f25_val = get_float_cell(df_sayfa1, 24, 5, default=2.4)
    g25_val = get_float_cell(df_sayfa1, 24, 6, default=10.0)
    h25_val = get_float_cell(df_sayfa1, 24, 7, default=3.0)
    i25_val = get_float_cell(df_sayfa1, 24, 8, default=10.0)

    # Kiremit: Row 28 (idx 27) -> G28=Col6 (m2), H28=Col7 (toplam kg)
    cati_alan_m2 = get_float_cell(df_sayfa1, 27, 6, default=200.0)
    kiremit_toplam_kg = get_float_cell(df_sayfa1, 27, 7, default=9000.0)

    # Seramik: Row 30 (idx 29) Tezgah, Row 32 (idx 31) Adet
    seramik_tezgah_kg = get_float_cell(df_sayfa1, 29, 8, default=44.1)
    seramik_adet_excel = get_float_cell(df_sayfa1, 31, 6, default=600.0) # G32
    seramik_mavi_kg = get_float_cell(df_sayfa1, 31, 7, default=2400.0)   # H32
    j32_val = get_float_cell(df_sayfa1, 31, 9, default=2444.1)           # J32 (44.1 + 2400)

    # Demir / Metaller: Row 49 (idx 48) Temel, Row 51 (idx 50) Katlar, Row 52 (idx 51) Toplam
    demir_temel_m2 = get_float_cell(df_sayfa1, 48, 6, default=200.0)
    demir_temel_toplam = get_float_cell(df_sayfa1, 48, 7, default=8000.0)
    demir_kat_m2 = get_float_cell(df_sayfa1, 50, 6, default=200.0)
    demir_kat_sayisi = get_float_cell(df_sayfa1, 50, 7, default=4.0)
    demir_kat_toplam = get_float_cell(df_sayfa1, 50, 8, default=16000.0)
    karisik_metal_kg = get_float_cell(df_sayfa1, 51, 8, default=24000.0)

    # Kağıt / Karton: Row 45 (idx 44)
    isci_sayisi = get_float_cell(df_sayfa1, 44, 6, default=4.0)
    santiye_suresi = get_float_cell(df_sayfa1, 44, 7, default=5.0)
    kagit_toplam_kg = get_float_cell(df_sayfa1, 44, 8, default=12.0)

    # Plastik: Row 21 (idx 20)
    plastik_pencere_adet = get_float_cell(df_sayfa1, 20, 8, default=6.0)
    plastik_daire_adet = get_float_cell(df_sayfa1, 20, 9, default=10.0)
    plastik_toplam_kg = get_float_cell(df_sayfa1, 20, 10, default=9.0)

    # Cam: Row 47 (idx 46)
    cam_daire_adet = get_float_cell(df_sayfa1, 46, 7, default=10.0)
    cam_miktari = get_float_cell(df_sayfa1, 46, 8, default=600.0)

    # Sıvalı Duvar: Row 11 (idx 10)
    siva_toplg = get_float_cell(df_sayfa1, 10, 9, default=66000.0)

    # --- Sayfa 2 Okumaları ---
    # H15 Hücresi (Row 15, Col H = idx 14, col 7): Tuğla + Beton + Kiremit + Seramik Toplamı
    h15_val = get_float_cell(df_sayfa2, 14, 7, default=330764.1)
    
    # Genel Toplam (Row 27, Col G = idx 26, col 6)
    genel_toplam_miktar = get_float_cell(df_sayfa2, 26, 6, default=422105.1)

    # Esenyurt Şablonu için 17 01 07 kodunun karşılığı DOĞRUDAN H15 (Sayfa2) hücresidir
    karisim_toplam_kg = h15_val if has_esenyurt_karisim else siva_toplg

    ahsap_toplg = f25_val * g25_val * h25_val * i25_val

    res = {
        "alan_m2": format_num(alan_m2),
        "kat_sayisi": format_num(kat_sayisi, 0),
        "cati_alan_m2": format_num(cati_alan_m2),
        "beton_toplam_kg": format_num(beton_toplg),
        "kiremit_toplam_kg": format_num(kiremit_toplam_kg),
        "seramik_tezgah_kg": format_num(seramik_tezgah_kg),
        "seramik_adet": format_num(seramik_adet_excel, 0),
        "g32": format_num(seramik_adet_excel, 0),
        "seramik_adet_toplam_kg": format_num(seramik_mavi_kg),
        "h32": format_num(seramik_mavi_kg),
        "seramik_genel_toplam_kg": format_num(j32_val),
        "j32": format_num(j32_val),
        "f25": format_num(f25_val),
        "g25": format_num(g25_val, 0),
        "h25": format_num(h25_val, 0),
        "i25": format_num(i25_val, 0),
        "k26": format_num(ahsap_toplg),
        "ahsap_toplam_kg": format_num(ahsap_toplg),
        "demir_temel_m2": format_num(demir_temel_m2),
        "demir_temel_toplam": format_num(demir_temel_toplam),
        "demir_kat_m2": format_num(demir_kat_m2),
        "demir_kat_sayisi": format_num(demir_kat_sayisi, 0),
        "demir_kat_toplam": format_num(demir_kat_toplam),
        "toplam_karisik_metal": format_num(karisik_metal_kg),
        "isci_sayisi": format_num(isci_sayisi, 0),
        "santiye_suresi": format_num(santiye_suresi, 0),
        "kagit_toplam_kg": format_num(kagit_toplam_kg),
        "plastik_pencere_adet": format_num(plastik_pencere_adet, 0),
        "plastik_daire_adet": format_num(plastik_daire_adet, 0),
        "plastik_toplam_kg": format_num(plastik_toplam_kg),
        "cam_daire_adet": format_num(cam_daire_adet, 0),
        "cam_miktari": format_num(cam_miktari),
        "siva_toplam_kg": format_num(siva_toplg),
        "karisim_toplam_kg": karisim_toplam_kg,
        "karisim_toplam_kg_fmt": format_num(karisim_toplam_kg),
        "17_01_07": format_num(karisim_toplam_kg),
        "genel_toplam_miktar": format_num(genel_toplam_miktar),
        "beton_toplam_ton": format_num(beton_toplg / 1000.0, 3),
        "kiremit_seramik_toplam_ton": format_num((kiremit_toplam_kg + j32_val) / 1000.0, 3),
        "ahsap_toplam_ton": format_num(ahsap_toplg / 1000.0, 3),
        "toplam_karisik_metal_ton": format_num(karisik_metal_kg / 1000.0, 3),
        "cam_miktari_ton": format_num(cam_miktari / 1000.0, 3),
        "genel_toplam_miktar_ton": format_num(genel_toplam_miktar / 1000.0, 3),
    }
    return res


def render_ayp_module():
    st.subheader("♻️ Atık Yönetim Planı (AYP) Rapor Oluşturucu")

    st.markdown("### 📑 AYP Rapor Şablonu ve Personel Seçimi")

    muhendisler_verisi = muhendisleri_excelden_oku()

    col_sablon, col_muh = st.columns(2)
    with col_sablon:
        secilen_sablon = st.selectbox(
            "Kullanılacak AYP Şablonunu Belirleyin:",
            options=list(SABLON_AYARLARI.keys()),
            key="ayp_sablon_secimi",
        )
    with col_muh:
        secilen_muhendis = st.selectbox(
            "Raporu Hazırlayan Çevre Mühendisi:",
            options=list(muhendisler_verisi.keys()),
            key="ayp_muhendis_secimi",
        )

    cfg = SABLON_AYARLARI[secilen_sablon]
    aktif_sablon_dosyasi = cfg["file_name"]
    excel_beklenen_label = cfg["label"]
    requires_excel = cfg["requires_excel"]

    st.markdown("---")

    # ==========================================
    # DİNAMİK SATIRLI TOPLU ZIP ÜRETİM ALANI (AYP)
    # ==========================================
    with st.expander("📦 Proje Bazlı Dinamik Çoklu AYP Raporu Üret (ZIP İndir)"):
        st.info("💡 Üretmek istediğiniz rapor sayısını belirleyin, her proje için gerekli bilgileri doldurun ve dosyaları yükleyin.")
        
        proje_sayisi = st.number_input("Kaç Adet AYP Raporu Üretilecek?", min_value=1, max_value=20, value=1, step=1, key="ayp_proje_sayisi")
        
        toplu_veri_listesi = []
        for i in range(int(proje_sayisi)):
            st.markdown(f"**--- Proje / Rapor {i+1} ---**")
            
            p_yapi_alani = 465.0
            p_kat_sayisi = 5
            p_cam_durumu = "Var"

            if cfg["is_sultanbeyli"] or cfg["is_sultangazi"]:
                t_upl = st.file_uploader(f"Proje {i+1} - Tutanak Excel", type=["xlsx", "xls"], key=f"tutanak_p_{i}")
                e_upl = "not_required"
                
                col_p1, col_p2, col_p3 = st.columns(3)
                with col_p1:
                    p_yapi_alani = st.number_input(f"Proje {i+1} - Toplam Yapı Alanı (m²):", min_value=0.0, value=465.0, step=10.0, key=f"yapi_alani_p_{i}")
                with col_p2:
                    p_kat_sayisi = st.number_input(f"Proje {i+1} - Kat Sayısı:", min_value=0, value=5, step=1, key=f"kat_sayisi_p_{i}")
                with col_p3:
                    p_cam_durumu = st.radio(f"Proje {i+1} - Cam Var mı?", options=["Var", "Yok"], index=0, horizontal=True, key=f"cam_durumu_p_{i}")
            else:
                p_col1, p_col2 = st.columns(2)
                with p_col1:
                    t_upl = st.file_uploader(f"Proje {i+1} - Tutanak Excel", type=["xlsx", "xls"], key=f"tutanak_p_{i}")
                with p_col2:
                    e_upl = st.file_uploader(f"Proje {i+1} - Hesaplama Excel", type=["xlsx", "xls"], key=f"excel_p_{i}")

            toplu_veri_listesi.append({
                "tutanak": t_upl,
                "excel": e_upl,
                "yapi_alani": p_yapi_alani,
                "kat_sayisi": p_kat_sayisi,
                "cam_durumu": p_cam_durumu
            })

        if st.button("🚀 Tüm Dinamik AYP Raporlarını ZIP Olarak Hazırla", key="btn_dinamik_toplu_ayp"):
            if secilen_muhendis == "Seçiniz...":
                st.warning("⚠️ Lütfen önce raporu hazırlayan çevre mühendisini seçin.")
            else:
                eksik_var_mi = any(item["tutanak"] is None or (requires_excel and item["excel"] is None) for item in toplu_veri_listesi)
                if eksik_var_mi:
                    st.error("⚠️ Lütfen tüm projeler için gerekli dosyaları eksiksiz yükleyin!")
                else:
                    try:
                        zip_buffer = io.BytesIO()
                        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                        template_path = os.path.join(base_dir, "templates", aktif_sablon_dosyasi)

                        if os.path.exists(template_path):
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                for idx, item in enumerate(toplu_veri_listesi):
                                    t_file = item["tutanak"]
                                    e_file = item["excel"]

                                    t_path = os.path.join(UPLOAD_FOLDER, f"p{idx}_{t_file.name}")
                                    with open(t_path, "wb") as f:
                                        f.write(t_file.getbuffer())

                                    if requires_excel:
                                        e_path = os.path.join(UPLOAD_FOLDER, f"p{idx}_{e_file.name}")
                                        with open(e_path, "wb") as f:
                                            f.write(e_file.getbuffer())

                                    raw_info = read_tutanak_details(t_path)
                                    info = {}
                                    if isinstance(raw_info, tuple):
                                        if len(raw_info) > 0 and isinstance(raw_info[0], dict):
                                            info = raw_info[0].copy()
                                    elif isinstance(raw_info, dict):
                                        info = raw_info.copy()

                                    adres_metni = info.get("adres", "")
                                    mahalle_match = re.search(
                                        r"([\w\sçğıöşüÇĞİÖŞÜ]+(?:Mahallesi|Mah\.?))",
                                        adres_metni,
                                        re.IGNORECASE,
                                    )
                                    bulunan_mahalle = mahalle_match.group(1).strip() if mahalle_match else "Belirtilmemiş"

                                    raw_tarih = info.get("tarih", info.get("tutanak_tarihi", datetime.now().strftime("%d.%m.%Y")))
                                    final_tarih = str(raw_tarih).strip() if raw_tarih and str(raw_tarih).strip() != "" else datetime.now().strftime("%d.%m.%Y")
                                    
                                    info.update({
                                        "mahalle": bulunan_mahalle,
                                        "tarih": final_tarih,
                                        "tutanak_tarihi": final_tarih,
                                        "rapor_tarihi": final_tarih,
                                        "bugun_tarihi": datetime.now().strftime("%d.%m.%Y"),
                                        "cevre_muhendisi": secilen_muhendis,
                                        "oda_sicil_no": muhendisler_verisi[secilen_muhendis]["sicil"],
                                        "tc_kimlik_no": muhendisler_verisi[secilen_muhendis]["tc"],
                                    })

                                    if cfg["is_sultanbeyli"]:
                                        toplam_yapi_alani_m2 = float(item["yapi_alani"])
                                        kat_sayisi = int(item["kat_sayisi"])
                                        cam_kat_sayisi = kat_sayisi if item["cam_durumu"] == "Var" else 0

                                        beton_toplam_kg = 0.25 * 1000.0 * toplam_yapi_alani_m2
                                        beton_toplam_ton = 0.25 * toplam_yapi_alani_m2
                                        tugla_kiremit_seramik_toplam_kg = 23.0 * toplam_yapi_alani_m2
                                        tugla_kiremit_seramik_toplam_ton = tugla_kiremit_seramik_toplam_kg / 1000.0
                                        cam_miktari_kg = 20.0 * 10.0 * cam_kat_sayisi
                                        cam_miktari_ton = cam_miktari_kg / 1000.0
                                        plastik_toplam_kg = 10.0 * cam_kat_sayisi
                                        plastik_toplam_ton = plastik_toplam_kg / 1000.0
                                        toplam_karisik_metal_kg = 40.0 * toplam_yapi_alani_m2
                                        toplam_karisik_metal_ton = toplam_karisik_metal_kg / 1000.0
                                        kablo_toplam_kg = 50.0 * kat_sayisi
                                        kablo_toplam_ton = kablo_toplam_kg / 1000.0
                                        kagit_toplam_kg = 23.0
                                        kagit_toplam_ton = 0.023

                                        info.update({
                                            "toplam_yapi_alani_m2": format_num(toplam_yapi_alani_m2),
                                            "kat_sayisi": str(kat_sayisi),
                                            "cam_kat_sayisi": str(cam_kat_sayisi),
                                            "beton_toplam_kg": format_num(beton_toplam_kg),
                                            "beton_toplam_ton": format_num(beton_toplam_ton, 1),
                                            "tugla_kiremit_seramik_toplam_kg": format_num(tugla_kiremit_seramik_toplam_kg),
                                            "tugla_kiremit_seramik_toplam_ton": format_num(tugla_kiremit_seramik_toplam_ton, 1),
                                            "cam_miktari_kg": format_num(cam_miktari_kg, 0),
                                            "cam_miktari_ton": format_num(cam_miktari_ton, 1),
                                            "plastik_toplam_kg": format_num(plastik_toplam_kg, 0),
                                            "plastik_toplam_ton": format_num(plastik_toplam_ton, 1),
                                            "toplam_karisik_metal_kg": format_num(toplam_karisik_metal_kg),
                                            "toplam_karisik_metal_ton": format_num(toplam_karisik_metal_ton, 1),
                                            "kablo_toplam_kg": format_num(kablo_toplam_kg, 0),
                                            "kablo_toplam_ton": format_num(kablo_toplam_ton, 1),
                                            "kagit_toplam_kg": format_num(kagit_toplam_kg, 0),
                                            "kagit_toplam_ton": format_num(kagit_toplam_ton, 3),
                                        })

                                    elif cfg["is_sultangazi"]:
                                        toplam_yapi_alani_m2 = float(item["yapi_alani"])
                                        kat_sayisi = int(item["kat_sayisi"])
                                        cam_kat_sayisi = kat_sayisi if item["cam_durumu"] == "Var" else 0

                                        beton_toplam_kg = 0.25 * 1000.0 * toplam_yapi_alani_m2
                                        beton_toplam_ton = 0.25 * toplam_yapi_alani_m2
                                        kiremit_seramik_toplam_kg = 13.0 * toplam_yapi_alani_m2
                                        kiremit_seramik_toplam_ton = kiremit_seramik_toplam_kg / 1000.0
                                        tugla_toplam_kg = 24.0 * toplam_yapi_alani_m2
                                        tugla_toplam_ton = tugla_toplam_kg / 1000.0
                                        cam_miktari_kg = 20.0 * 10.0 * cam_kat_sayisi
                                        cam_miktari_ton = cam_miktari_kg / 1000.0
                                        plastik_toplam_kg = 10.0 * cam_kat_sayisi
                                        plastik_toplam_ton = plastik_toplam_kg / 1000.0
                                        toplam_karisik_metal_kg = 40.0 * toplam_yapi_alani_m2
                                        toplam_karisik_metal_ton = toplam_karisik_metal_kg / 1000.0
                                        kablo_toplam_kg = 50.0 * kat_sayisi
                                        kablo_toplam_ton = kablo_toplam_kg / 1000.0

                                        info.update({
                                            "toplam_yapi_alani_m2": format_num(toplam_yapi_alani_m2),
                                            "kat_sayisi": str(kat_sayisi),
                                            "cam_kat_sayisi": str(cam_kat_sayisi),
                                            "beton_toplam_kg": format_num(beton_toplam_kg),
                                            "kiremit_seramik_toplam_kg": format_num(kiremit_seramik_toplam_kg),
                                            "kiremit_seramik_toplam_ton": format_num(kiremit_seramik_toplam_ton, 1),
                                            "tugla_toplam_kg": format_num(tugla_toplam_kg),
                                            "tugla_toplam_ton": format_num(tugla_toplam_ton, 1),
                                            "cam_miktari_kg": format_num(cam_miktari_kg, 0),
                                            "cam_miktari_ton": format_num(cam_miktari_ton, 1),
                                            "plastik_toplam_kg": format_num(plastik_toplam_kg, 0),
                                            "plastik_toplam_ton": format_num(plastik_toplam_ton, 1),
                                            "toplam_karisik_metal_kg": format_num(toplam_karisik_metal_kg),
                                            "toplam_karisik_metal_ton": format_num(toplam_karisik_metal_ton, 1),
                                            "kablo_toplam_kg": format_num(kablo_toplam_kg, 0),
                                            "kablo_toplam_ton": format_num(kablo_toplam_ton, 1),
                                        })

                                    else:
                                        excel_info = process_ayp_excel(e_path, has_esenyurt_karisim=cfg["has_esenyurt_karisim"])
                                        info.update(excel_info)

                                    doc = DocxTemplate(template_path)
                                    render_context = sanitize_context_for_jinja(info)
                                    doc.render(render_context)

                                    musteri_adi = render_context.get("musteri_adi", f"Proje_{idx+1}")
                                    safe_name = "".join(c for c in str(musteri_adi) if c.isalnum() or c in (' ', '_', '-')).strip()

                                    out_io = io.BytesIO()
                                    doc.save(out_io)
                                    zf.writestr(f"AYP_Raporu_{safe_name}.docx", out_io.getvalue())

                            zip_buffer.seek(0)
                            st.success("✅ Tüm dinamik AYP raporları başarıyla paketlendi!")
                            st.download_button(
                                label="⬇️ Toplu AYP Raporları Arşivini İndir (.zip)",
                                data=zip_buffer.getvalue(),
                                file_name="Toplu_AYP_Raporlari.zip",
                                mime="application/zip",
                                key="dl_toplu_ayp_zip"
                            )
                        else:
                            st.error(f"❌ Şablon bulunamadı: {aktif_sablon_dosyasi}")
                    except Exception as e:
                        st.error(f"Toplu AYP üretiminde hata: {e}")

    st.markdown("---")

    # ==========================================
    # TEKLİ RAPOR OLUŞTURMA ALANI
    # ==========================================
    toplam_yapi_alani_input = 0.0
    kat_sayisi_input = 0
    cam_durumu_input = "Var"

    if cfg["is_sultangazi"] or cfg["is_sultanbeyli"]:
        ilce_adi = "Sultanbeyli" if cfg["is_sultanbeyli"] else "Sultangazi"
        st.markdown(f"### 🏗️ Tekli Rapor için {ilce_adi} Yapı Bilgileri")
        col_sg1, col_sg2, col_sg3 = st.columns(3)
        with col_sg1:
            toplam_yapi_alani_input = st.number_input(
                "Toplam Yapı Alanı (m²):",
                min_value=0.0,
                value=465.0,
                step=10.0,
                key="yapi_alani_input_single",
            )
        with col_sg2:
            kat_sayisi_input = st.number_input(
                "Kat Sayısı:",
                min_value=0,
                value=5,
                step=1,
                key="kat_sayisi_input_single",
            )
        with col_sg3:
            cam_durumu_input = st.radio(
                "Cam Var mı?",
                options=["Var", "Yok"],
                index=0,
                horizontal=True,
                key="cam_durumu_input_single",
            )
        st.markdown("---")

    foto_file = None
    if requires_excel:
        col1, col2 = st.columns(2)
        with col1:
            tutanak_file = st.file_uploader(
                "📂 1. Tutanak Dosyası (Excel - Künye için):",
                type=["xlsx", "xls"],
                key="ayp_tutanak_single",
            )
        with col2:
            ayp_file = st.file_uploader(
                excel_beklenen_label,
                type=["xlsx", "xls"],
                key="ayp_excel_single",
            )
    else:
        ayp_file = None
        tutanak_file = st.file_uploader(
            "📂 1. Tutanak Dosyası (Excel - Künye için):",
            type=["xlsx", "xls"],
            key="ayp_tutanak_single",
        )
        st.info(
            "ℹ️ Seçilen şablon için hesaplama Excel'i gerekmez. Hesaplama yukarıdaki değerlere göre otomatik yapılacaktır."
        )

    if cfg.get("is_pendik") and cfg.get("pendik_tip") == 1:
        st.markdown("### 📸 Rapor Fotoğrafı")
        foto_file = st.file_uploader(
            "📂 Şablon - 1 için Fotoğraf Yükleyin:",
            type=["png", "jpg", "jpeg"],
            key="pendik_foto_1_single",
        )

    can_proceed = tutanak_file is not None and (
        ayp_file is not None or not requires_excel
    )

    if can_proceed:
        try:
            os.makedirs(UPLOAD_FOLDER, exist_ok=True)

            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())

            raw_info = read_tutanak_details(tutanak_path)
            info = {}
            if isinstance(raw_info, tuple):
                if len(raw_info) > 0 and isinstance(raw_info[0], dict):
                    info = raw_info[0].copy()
            elif isinstance(raw_info, dict):
                info = raw_info.copy()

            adres_metni = info.get("adres", "")
            mahalle_match = re.search(
                r"([\w\sçğıöşüÇĞİÖŞÜ]+(?:Mahallesi|Mah\.?))",
                adres_metni,
                re.IGNORECASE,
            )
            bulunan_mahalle = (
                mahalle_match.group(1).strip()
                if mahalle_match
                else "Belirtilmemiş"
            )

            raw_tarih = info.get(
                "tarih",
                info.get(
                    "tutanak_tarihi", datetime.now().strftime("%d.%m.%Y")
                ),
            )
            final_tarih = (
                str(raw_tarih).strip()
                if raw_tarih and str(raw_tarih).strip() != ""
                else datetime.now().strftime("%d.%m.%Y")
            )
            info.update({
                "mahalle": bulunan_mahalle,
                "tarih": final_tarih,
                "tutanak_tarihi": final_tarih,
                "rapor_tarihi": final_tarih,
                "bugun_tarihi": datetime.now().strftime("%d.%m.%Y"),
                "cevre_muhendisi": secilen_muhendis,
                "oda_sicil_no": muhendisler_verisi[secilen_muhendis]["sicil"],
                "tc_kimlik_no": muhendisler_verisi[secilen_muhendis]["tc"],
            })

            if cfg["is_sultanbeyli"]:
                toplam_yapi_alani_m2 = float(toplam_yapi_alani_input)
                kat_sayisi = int(kat_sayisi_input)
                cam_kat_sayisi = (
                    kat_sayisi if cam_durumu_input == "Var" else 0
                )

                beton_toplam_kg = 0.25 * 1000.0 * toplam_yapi_alani_m2
                beton_toplam_ton = 0.25 * toplam_yapi_alani_m2

                tugla_kiremit_seramik_toplam_kg = 23.0 * toplam_yapi_alani_m2
                tugla_kiremit_seramik_toplam_ton = (
                    tugla_kiremit_seramik_toplam_kg / 1000.0
                )

                cam_miktari_kg = 20.0 * 10.0 * cam_kat_sayisi
                cam_miktari_ton = cam_miktari_kg / 1000.0

                plastik_toplam_kg = 10.0 * cam_kat_sayisi
                plastik_toplam_ton = plastik_toplam_kg / 1000.0

                toplam_karisik_metal_kg = 40.0 * toplam_yapi_alani_m2
                toplam_karisik_metal_ton = toplam_karisik_metal_kg / 1000.0

                kablo_toplam_kg = 50.0 * kat_sayisi
                kablo_toplam_ton = kablo_toplam_kg / 1000.0

                kagit_toplam_kg = 23.0
                kagit_toplam_ton = 0.023

                info.update({
                    "toplam_yapi_alani_m2": format_num(toplam_yapi_alani_m2),
                    "kat_sayisi": str(kat_sayisi),
                    "cam_kat_sayisi": str(cam_kat_sayisi),
                    "beton_toplam_kg": format_num(beton_toplam_kg),
                    "beton_toplam_ton": format_num(beton_toplam_ton, 1),
                    "tugla_kiremit_seramik_toplam_kg": format_num(
                        tugla_kiremit_seramik_toplam_kg
                    ),
                    "tugla_kiremit_seramik_toplam_ton": format_num(
                        tugla_kiremit_seramik_toplam_ton, 1
                    ),
                    "cam_miktari_kg": format_num(cam_miktari_kg, 0),
                    "cam_miktari_ton": format_num(cam_miktari_ton, 1),
                    "plastik_toplam_kg": format_num(plastik_toplam_kg, 0),
                    "plastik_toplam_ton": format_num(plastik_toplam_ton, 1),
                    "toplam_karisik_metal_kg": format_num(
                        toplam_karisik_metal_kg
                    ),
                    "toplam_karisik_metal_ton": format_num(
                        toplam_karisik_metal_ton, 1
                    ),
                    "kablo_toplam_kg": format_num(kablo_toplam_kg, 0),
                    "kablo_toplam_ton": format_num(kablo_toplam_ton, 1),
                    "kagit_toplam_kg": format_num(kagit_toplam_kg, 0),
                    "kagit_toplam_ton": format_num(kagit_toplam_ton, 3),
                })

            elif cfg["is_sultangazi"]:
                toplam_yapi_alani_m2 = float(toplam_yapi_alani_input)
                kat_sayisi = int(kat_sayisi_input)
                cam_kat_sayisi = (
                    kat_sayisi if cam_durumu_input == "Var" else 0
                )

                beton_toplam_kg = 0.25 * 1000.0 * toplam_yapi_alani_m2
                beton_toplam_ton = 0.25 * toplam_yapi_alani_m2

                kiremit_seramik_toplam_kg = 13.0 * toplam_yapi_alani_m2
                kiremit_seramik_toplam_ton = kiremit_seramik_toplam_kg / 1000.0

                tugla_toplam_kg = 24.0 * toplam_yapi_alani_m2
                tugla_toplam_ton = tugla_toplam_kg / 1000.0

                cam_miktari_kg = 20.0 * 10.0 * cam_kat_sayisi
                cam_miktari_ton = cam_miktari_kg / 1000.0

                plastik_toplam_kg = 10.0 * cam_kat_sayisi
                plastik_toplam_ton = plastik_toplam_kg / 1000.0

                toplam_karisik_metal_kg = 40.0 * toplam_yapi_alani_m2
                toplam_karisik_metal_ton = toplam_karisik_metal_kg / 1000.0

                kablo_toplam_kg = 50.0 * kat_sayisi
                kablo_toplam_ton = kablo_toplam_kg / 1000.0

                info.update({
                    "toplam_yapi_alani_m2": format_num(toplam_yapi_alani_m2),
                    "kat_sayisi": str(kat_sayisi),
                    "cam_kat_sayisi": str(cam_kat_sayisi),
                    "beton_toplam_kg": format_num(beton_toplam_kg),
                    "kiremit_seramik_toplam_kg": format_num(
                        kiremit_seramik_toplam_kg
                    ),
                    "kiremit_seramik_toplam_ton": format_num(
                        kiremit_seramik_toplam_ton, 1
                    ),
                    "tugla_toplam_kg": format_num(tugla_toplam_kg),
                    "tugla_toplam_ton": format_num(tugla_toplam_ton, 1),
                    "cam_miktari_kg": format_num(cam_miktari_kg, 0),
                    "cam_miktari_ton": format_num(cam_miktari_ton, 1),
                    "plastik_toplam_kg": format_num(plastik_toplam_kg, 0),
                    "plastik_toplam_ton": format_num(plastik_toplam_ton, 1),
                    "toplam_karisik_metal_kg": format_num(
                        toplam_karisik_metal_kg
                    ),
                    "toplam_karisik_metal_ton": format_num(
                        toplam_karisik_metal_ton, 1
                    ),
                    "kablo_toplam_kg": format_num(kablo_toplam_kg, 0),
                    "kablo_toplam_ton": format_num(kablo_toplam_ton, 1),
                })

            elif requires_excel and ayp_file is not None:
                ayp_path = os.path.join(UPLOAD_FOLDER, ayp_file.name)
                with open(ayp_path, "wb") as f:
                    f.write(ayp_file.getbuffer())

                excel_info = process_ayp_excel(ayp_path, has_esenyurt_karisim=cfg["has_esenyurt_karisim"])
                info.update(excel_info)

            render_context = sanitize_context_for_jinja(info)

            st.success(
                f"✅ '{secilen_sablon}' için değişkenler başarıyla hazırlandı."
            )

            if st.button("🚀 AYP Raporunu Oluştur", type="primary"):
                if secilen_muhendis == "Seçiniz...":
                    st.warning("⚠️ Lütfen raporu hazırlayan çevre mühendisini seçin.")
                    return

                current_script_dir = os.path.dirname(
                    os.path.abspath(__file__)
                )
                possible_paths = [
                    os.path.join(
                        current_script_dir, "templates", aktif_sablon_dosyasi
                    ),
                    os.path.join(current_script_dir, aktif_sablon_dosyasi),
                    os.path.join(
                        os.getcwd(), "templates", aktif_sablon_dosyasi
                    ),
                    os.path.join(os.getcwd(), aktif_sablon_dosyasi),
                ]
                template_path = next(
                    (p for p in possible_paths if os.path.exists(p)), None
                )

                if template_path:
                    doc = DocxTemplate(template_path)

                    if cfg.get("is_pendik") and cfg.get("pendik_tip") == 1:
                        if foto_file is not None:
                            foto_path = os.path.join(
                                UPLOAD_FOLDER, foto_file.name
                            )
                            with open(foto_path, "wb") as f:
                                f.write(foto_file.getbuffer())

                            try:
                                img = Image.open(foto_path)
                                img = ImageOps.exif_transpose(img)
                                img.save(foto_path)
                            except Exception:
                                pass

                            render_context["foto"] = InlineImage(
                                doc, foto_path, width=Inches(4.0)
                            )
                        else:
                            render_context["foto"] = ""

                    doc.render(render_context)

                    musteri_adi = render_context.get("musteri_adi", "Musteri")
                    safe_musteri_adi = re.sub(
                        r'[\\/*?:"<>|]', "_", str(musteri_adi)
                    )
                    output_filename = f"AYP_Raporu_{safe_musteri_adi}.docx"
                    output_path = os.path.join(UPLOAD_FOLDER, output_filename)
                    doc.save(output_path)

                    with open(output_path, "rb") as f:
                        file_data = f.read()

                    st.download_button(
                        label="📥 AYP Raporunu İndir (.docx)",
                        data=file_data,
                        file_name=output_filename,
                        mime=(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        ),
                    )
                else:
                    st.error(
                        f"❌ Şablon bulunamadı: '{aktif_sablon_dosyasi}'"
                    )

        except Exception as e:
            st.error(f"❌ İşlem sırasında hata: {e}")

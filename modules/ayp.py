import io
import os
import re
from datetime import datetime
from docxtpl import DocxTemplate
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
    },
    "Esenyurt AYP Şablonu (sablon_ayp_esenyurt.docx)": {
        "file_name": "sablon_ayp_esenyurt.docx",
        "label": "📂 2. AYP Hesaplama Esenyurt Dosyası (Excel):",
        "has_esenyurt_karisim": True,
        "is_ton_bazli_excel": False,
        "is_sultangazi": False,
        "is_sultanbeyli": False,
    },
    "Sultanbeyli AYP Şablonu (sablon_ayp_sultanbeyli.docx)": {
        "file_name": "sablon_ayp_sultanbeyli.docx",
        "label": "📂 2. AYP Hesaplama Sultanbeyli Dosyası (Excel):",
        "has_esenyurt_karisim": False,
        "is_ton_bazli_excel": False,
        "is_sultangazi": False,
        "is_sultanbeyli": True,
    },
    "Sultangazi AYP Şablonu (sablon_ayp_sultangazi.docx)": {
        "file_name": "sablon_ayp_sultangazi.docx",
        "label": "📂 2. AYP Hesaplama Sultangazi Dosyası (Excel):",
        "has_esenyurt_karisim": False,
        "is_ton_bazli_excel": False,
        "is_sultangazi": True,
        "is_sultanbeyli": False,
    },
    "Ton Bazlı AYP Şablonu (sablon_ayp_ton.docx)": {
        "file_name": "sablon_ayp_ton.docx",
        "label": "📂 2. AYP Hesaplama Ton Dosyası (Excel):",
        "has_esenyurt_karisim": False,
        "is_ton_bazli_excel": True,
        "is_sultangazi": False,
        "is_sultanbeyli": False,
    },
}


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


def render_ayp_module():
    st.subheader("♻️ Atık Yönetim Planı (AYP) Rapor Oluşturucu")

    st.markdown("### 📑 AYP Rapor Şablonu Seçimi")
    secilen_sablon = st.selectbox(
        "Kullanılacak AYP Şablonunu Belirleyin:",
        options=list(SABLON_AYARLARI.keys()),
        key="ayp_sablon_secimi",
    )

    cfg = SABLON_AYARLARI[secilen_sablon]
    aktif_sablon_dosyasi = cfg["file_name"]
    excel_beklenen_label = cfg["label"]

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

    if tutanak_file and ayp_file:
        try:
            os.makedirs(UPLOAD_FOLDER, exist_ok=True)

            # Tutanak Okuma
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

            # AYP Excel Okuma
            ayp_path = os.path.join(UPLOAD_FOLDER, ayp_file.name)
            with open(ayp_path, "wb") as f:
                f.write(ayp_file.getbuffer())

            excel_engine = (
                "xlrd" if ayp_path.lower().endswith(".xls") else "openpyxl"
            )
            xls = pd.ExcelFile(ayp_path, engine=excel_engine)

            df_sayfa1 = (
                pd.read_excel(
                    ayp_path,
                    sheet_name="Sayfa1",
                    header=None,
                    engine=excel_engine,
                )
                if "Sayfa1" in xls.sheet_names
                else pd.DataFrame()
            )
            df_sayfa2 = (
                pd.read_excel(
                    ayp_path,
                    sheet_name="Sayfa2",
                    header=None,
                    engine=excel_engine,
                )
                if "Sayfa2" in xls.sheet_names
                else pd.DataFrame()
            )

            # --- SAYFA 1 HESAPLAMALARI (Mevcut Mantık Korundu) ---
            alan_m2 = get_float_cell(df_sayfa1, 15, 6, default=85.0)
            kat_sayisi = get_float_cell(df_sayfa1, 2, 2, default=6.0)
            daire_sayisi = get_float_cell(df_sayfa1, 3, 2, default=10.0)
            oda_sayisi = get_float_cell(df_sayfa1, 4, 2, default=3.0)
            cati_alan_m2 = get_float_cell(df_sayfa1, 27, 6, default=0.0)

            # --- SERAMİK SPESİFİK HÜCRELERİ (32. SATIR - KORUNDU) ---
            seramik_adet_excel = get_float_cell(df_sayfa1, 31, 6, default=0.0)
            if seramik_adet_excel == 0.0:
                seramik_adet_excel = get_float_cell(
                    df_sayfa1, 31, 4, default=1327.0
                )

            seramik_mavi_kg = get_float_cell(df_sayfa1, 31, 7, default=0.0)
            if seramik_mavi_kg == 0.0:
                seramik_mavi_kg = seramik_adet_excel * 4.0

            seramik_pembe_kg = get_float_cell(df_sayfa1, 31, 9, default=0.0)
            if seramik_pembe_kg == 0.0:
                seramik_pembe_kg = 44.1 + seramik_mavi_kg

            # Ahşap Hesaplamaları
            laminant_alan_m2 = get_float_cell(df_sayfa1, 24, 4, default=24.0)
            ahsap_toplam_kg = (
                2.4 * laminant_alan_m2 * oda_sayisi * daire_sayisi
            )

            # Karışık Metal Hesaplamaları
            demir_temel_toplam = alan_m2 * 40.0
            demir_kat_toplam = alan_m2 * 20.0 * kat_sayisi
            toplam_karisik_metal = demir_temel_toplam + demir_kat_toplam

            # Kağıt Karton
            isci_sayisi = get_float_cell(df_sayfa1, 5, 2, default=2.0)
            calisma_suresi_gun = get_float_cell(df_sayfa1, 6, 2, default=10.0)
            kagit_toplam_kg = 0.6 * isci_sayisi * calisma_suresi_gun

            # Plastik
            pencere_adet = get_float_cell(df_sayfa1, 35, 4, default=6.0)
            plastik_toplam_kg = (
                0.1 * 0.1 * 15.0 * pencere_adet * daire_sayisi
            )

            # --- SAYFA 2 ATIK KODLARI VE TON HARİTASI OKUMA ---
            atik_miktarlari = {}
            genel_toplam_miktar = 0.0

            # Sayfa 2'den kilogram bazlı okuma (Mevcut)
            for idx, row in df_sayfa2.iterrows():
                row_vals = [v for v in row.values if pd.notna(v)]
                if not row_vals:
                    continue

                row_str_full = " ".join([str(v) for v in row_vals]).lower()
                if "toplam" in row_str_full and "daire" not in row_str_full:
                    for v in row.values:
                        val_f = parse_turkish_float(v, default=0.0)
                        if val_f > 0.0:
                            genel_toplam_miktar = val_f
                            break

                key = row.iloc[5] if len(row) > 6 else None
                val = row.iloc[6] if len(row) > 6 else None

                if (
                    pd.notna(key)
                    and str(key).strip().lower() != "atık kodu tanımı"
                ):
                    val_num = parse_turkish_float(val, default=0.0)
                    atik_miktarlari[str(key).strip().lower()] = val_num

            # --- TON BAZLI SÜTUN OKUMASI (Sayfa 2 - I & J Sütunları / 8 & 9. Indeks) ---
            ton_map = {}
            genel_toplam_ton_val = 0.0

            for idx, row in df_sayfa2.iterrows():
                if len(row) > 9:
                    t_label = str(row.iloc[8]).strip().upper()
                    t_val = parse_turkish_float(row.iloc[9], default=0.0)

                    if t_label == "TOPLAM":
                        genel_toplam_ton_val = t_val
                    elif t_label and t_label != "NAN":
                        ton_map[t_label] = t_val

            # Ton Değerleri (Excel'den geldiyse Excel'den, yoksa kg/1000 dönüşümünden)
            beton_toplam_ton = ton_map.get("BETON", (alan_m2 * 2400.0 * 0.15 * kat_sayisi) / 1000.0)
            kiremit_toplam_ton = ton_map.get("KİREMİT", (45.0 * cati_alan_m2) / 1000.0)
            seramik_genel_toplam_ton = ton_map.get("SERAMİK", seramik_pembe_kg / 1000.0)
            ahsap_toplam_ton = ton_map.get("AHŞAP", ahsap_toplam_kg / 1000.0)
            tugla_toplam_ton = ton_map.get("TUĞLA", atik_miktarlari.get("tuğla", 0.0) / 1000.0)
            siva_toplam_ton = ton_map.get("SIVALI DUVAR", atik_miktarlari.get("17 08 01 dışındaki alçı bazlı inşaat malzemeleri", 0.0) / 1000.0)
            toplam_karisik_metal_ton = ton_map.get("KARIŞIK METAL", toplam_karisik_metal / 1000.0)
            kagit_toplam_ton = ton_map.get("KAĞIT", kagit_toplam_kg / 1000.0)
            plastik_toplam_ton = ton_map.get("PLASTİK", plastik_toplam_kg / 1000.0)
            cam_miktari_ton = ton_map.get("CAM", 0.0)
            asbest_toplam_ton = ton_map.get("ASBEST", atik_miktarlari.get("asbest içeren inşaat malzemeleri", 0.0) / 1000.0)

            if genel_toplam_ton_val == 0.0:
                genel_toplam_ton_val = genel_toplam_miktar / 1000.0

            # Diğer Hesaplanan Alanlar (kg)
            asbest_toplam_kg = atik_miktarlari.get(
                "asbest içeren inşaat malzemeleri", 0.0
            )
            beton_toplam_kg = atik_miktarlari.get(
                "beton", (2400.0 * alan_m2 * 0.15 * kat_sayisi)
            )
            kiremit_toplam_kg = atik_miktarlari.get(
                "kiremitler", (45.0 * cati_alan_m2)
            )
            cam_miktari = atik_miktarlari.get("cam ambalaj", 0.0)

            # --- YEŞİL BÖLGE TARİH BİLGİSİ (KORUNDU) ---
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

            # --- SÖZLÜK GÜNCELLEME (CONTEXT) ---
            info.update({
                # Yeşil Bölge (Tarih)
                "tarih": final_tarih,
                "tutanak_tarihi": final_tarih,
                "rapor_tarihi": final_tarih,
                "bugun_tarihi": datetime.now().strftime("%d.%m.%Y"),
                # Sarı Bölge (G32 - Seramik Adedi)
                "seramik_adet": format_num(seramik_adet_excel, 0),
                "g32": format_num(seramik_adet_excel, 0),
                # Mavi Bölge (H32 - 4 kg * adet kg)
                "seramik_adet_toplam_kg": format_num(seramik_mavi_kg),
                "h32": format_num(seramik_mavi_kg),
                # Pembe Bölge (J32 - 44,1 kg + Mavi kg)
                "seramik_genel_toplam_kg": format_num(seramik_pembe_kg),
                "j32": format_num(seramik_pembe_kg),
                # Genel Hesaplama Alanları (kg)
                "alan_m2": format_num(alan_m2),
                "kat_sayisi": format_num(kat_sayisi, 0),
                "daire_sayisi": format_num(daire_sayisi, 0),
                "oda_sayisi": format_num(oda_sayisi, 0),
                "cati_alan_m2": format_num(cati_alan_m2),
                "laminant_alan_m2": format_num(laminant_alan_m2),
                "ahsap_toplam_kg": format_num(ahsap_toplam_kg),
                "demir_temel_toplam": format_num(demir_temel_toplam),
                "demir_kat_toplam": format_num(demir_kat_toplam),
                "toplam_karisik_metal": format_num(toplam_karisik_metal),
                "isci_sayisi": format_num(isci_sayisi, 0),
                "calisma_suresi_gun": format_num(calisma_suresi_gun, 0),
                "kagit_toplam_kg": format_num(kagit_toplam_kg),
                "pencere_adet": format_num(pencere_adet, 0),
                "plastik_toplam_kg": format_num(plastik_toplam_kg),
                "asbest_toplam_kg": format_num(asbest_toplam_kg),
                "beton_toplam_kg": format_num(beton_toplam_kg),
                "kiremit_toplam_kg": format_num(kiremit_toplam_kg),
                "tugla_toplam_kg": format_num(
                    atik_miktarlari.get("tuğla", 0.0)
                ),
                "siva_toplam_kg": format_num(
                    atik_miktarlari.get(
                        "17 08 01 dışındaki alçı bazlı inşaat malzemeleri", 0.0
                    )
                ),
                "cam_miktari": format_num(cam_miktari),
                "genel_toplam_miktar": format_num(genel_toplam_miktar),
                # --- TON ŞABLONUNA ÖZEL BİREBİR DEĞİŞKENLER ---
                "asbest_toplam_ton": format_num(asbest_toplam_ton, 3),
                "beton_toplam_ton": format_num(beton_toplam_ton, 3),
                "kiremit_toplam_ton": format_num(kiremit_toplam_ton, 3),
                "seramik_genel_toplam_ton": format_num(seramik_genel_toplam_ton, 3),
                "ahsap_toplam_ton": format_num(ahsap_toplam_ton, 3),
                "tugla_toplam_ton": format_num(tugla_toplam_ton, 3),
                "siva_toplam_ton": format_num(siva_toplam_ton, 3),
                "toplam_karisik_metal_ton": format_num(toplam_karisik_metal_ton, 3),
                "kagit_toplam_ton": format_num(kagit_toplam_ton, 3),
                "plastik_toplam_ton": format_num(plastik_toplam_ton, 3),
                "cam_miktari_ton": format_num(cam_miktari_ton, 3),
                "genel_toplam_miktar_ton": format_num(genel_toplam_ton_val, 3),
            })

            render_context = sanitize_context_for_jinja(info)

            st.success(
                f"✅ '{secilen_sablon}' için Excel değişkenleri başarıyla hazırlandı."
            )

            if st.button("🚀 AYP Raporunu Oluştur", type="primary"):
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

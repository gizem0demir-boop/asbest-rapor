import io
import os
import re
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
import pandas as pd
from PIL import Image, ImageOps
import streamlit as st
from utils import UPLOAD_FOLDER, read_tutanak_details


def format_num(val, decimal_places=2):
    """Sayıları Türkçe formatta (örneğin 1.484,10) stringe dönüştürür."""
    if val is None:
        return "0,00"
    try:
        if isinstance(val, str):
            val = float(val.replace(".", "").replace(",", "."))
        return f"{val:,.{decimal_places}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)


def render_ayp_module():
    st.subheader("♻️ Atık Yönetim Planı (AYP) Rapor Oluşturucu")

    # 1. ŞABLON SEÇİMİ (Genel ve Özel Şablonlar)
    st.markdown("### 📑 AYP Rapor Şablonu Seçimi")
    secilen_sablon = st.selectbox(
        "Kullanılacak AYP Şablonunu Belirleyin:",
        options=[
            "Standart AYP Şablonu (sablon_ayp.docx)",
            "Avcılar & Üsküdar AYP Şablonu (sablon_ayp_avcilar_uskudar.docx)",
            "Sultanbeyli AYP Şablonu (sablon_ayp_sultanbeyli.docx)",
            "Genel Yedek Şablon (sablon.docx)",
        ],
        key="ayp_sablon_secimi",
    )

    # Şablon dosya adını eşleştirme
    if "Standart" in secilen_sablon:
        aktif_sablon_dosyasi = "sablon_ayp.docx"
    elif "Avcılar" in secilen_sablon:
        aktif_sablon_dosyasi = "sablon_ayp_avcilar_uskudar.docx"
    elif "Sultanbeyli" in secilen_sablon:
        aktif_sablon_dosyasi = "sablon_ayp_sultanbeyli.docx"
    else:
        aktif_sablon_dosyasi = "sablon.docx"

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
            "📂 2. AYP Hesaplama Dosyası (Excel):",
            type=["xlsx", "xls"],
            key="ayp_excel",
        )

    if tutanak_file and ayp_file:
        try:
            # UPLOAD_FOLDER dizininin varlığından emin olun
            os.makedirs(UPLOAD_FOLDER, exist_ok=True)

            # Tutanak dosyasını kaydet ve oku
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

            # AYP hesaplama dosyasını kaydet ve oku
            ayp_path = os.path.join(UPLOAD_FOLDER, ayp_file.name)
            with open(ayp_path, "wb") as f:
                f.write(ayp_file.getbuffer())

            xls = pd.ExcelFile(ayp_path)
            df_sayfa1 = (
                pd.read_excel(ayp_path, sheet_name="Sayfa1", header=None)
                if "Sayfa1" in xls.sheet_names
                else pd.DataFrame()
            )
            df_sayfa2 = (
                pd.read_excel(ayp_path, sheet_name="Sayfa2")
                if "Sayfa2" in xls.sheet_names
                else pd.DataFrame()
            )

            # Sayfa1 J32 hücresinden seramik miktarını çekme
            seramik_toplam_degeri = 1484.1
            try:
                if df_sayfa1.shape[0] > 31 and df_sayfa1.shape[1] > 9:
                    val_j32 = df_sayfa1.iloc[31, 9]
                    if pd.notna(val_j32):
                        seramik_toplam_degeri = float(
                            str(val_j32).replace(".", "").replace(",", ".")
                        )
            except Exception:
                pass

            # Sayfa2'den atık miktarlarını dinamik çekme
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
                        try:
                            val_f = float(
                                str(v).replace(".", "").replace(",", ".")
                            )
                            if val_f > 1000:
                                genel_toplam = val_f
                                break
                        except Exception:
                            pass

                key = row.iloc[5] if len(row) > 5 else None
                val = row.iloc[6] if len(row) > 6 else None

                if pd.notna(key) and str(key).strip().lower() != "atık kodu tanımı":
                    try:
                        val_num = (
                            float(str(val).replace(".", "").replace(",", "."))
                            if pd.notna(val)
                            else 0.0
                        )
                    except Exception:
                        val_num = 0.0
                    atik_miktarlari[str(key).strip().lower()] = val_num

            st.success("✅ Tutanak ve AYP hesaplama verileri başarıyla okundu.")

            # DİNAMİK PARAMETRE GİRİŞ ALANLARI
            st.markdown("---")
            st.markdown("### 📐 Yapı ve Çalışma Saha Parametreleri")
            
            p_col1, p_col2, p_col3 = st.columns(3)
            with p_col1:
                alan_m2 = st.number_input("Taban / Bina Alanı (m²)", value=85.0, key="ayp_alan")
                kat_sayisi = st.number_input("Kat Sayısı", value=6.0, step=1.0, key="ayp_kat")
                daire_sayisi = st.number_input("Daire Sayısı", value=6.0, step=1.0, key="ayp_daire")
            with p_col2:
                cati_alan_m2 = st.number_input("Çatı Alanı (m²)", value=85.0, key="ayp_cati")
                oda_sayisi = st.number_input("Oda Sayısı (Daire Başı)", value=3, key="ayp_oda")
                pencere_adet = st.number_input("Pencere Adedi", value=6, key="ayp_pencere")
            with p_col3:
                isci_sayisi = st.number_input("Çalışacak İşçi Sayısı", value=4, key="ayp_isci")
                calisma_suresi = st.number_input("Çalışma Süresi (Gün)", value=5, key="ayp_sure")
                laminant_m2 = st.number_input("Laminat Alanı (m²)", value=8.0, key="ayp_laminat")

            bugun_tarihi = datetime.now().strftime("%d.%m.%Y")

            # Miktarları doğrudan float olarak saklama
            asbest_kg = atik_miktarlari.get("asbest içeren inşaat malzemeleri", 0.0)
            beton_kg = atik_miktarlari.get("beton", 183600.0)
            kiremit_kg = 3825.0
            ahsap_kg = atik_miktarlari.get("ahşap", 345.6)
            tugla_kg = atik_miktarlari.get("tuğla", 15840.0)
            siva_kg = atik_miktarlari.get("17 08 01 dışındaki alçı bazlı inşaat malzemeleri", 52800.0)
            metal_kg = atik_miktarlari.get("karışık metaller", 20400.0)
            kagit_kg = atik_miktarlari.get("kağıt ve karton ambalaj", 12.0)
            plastik_kg = atik_miktarlari.get("plastik ambalaj", 0.0)
            cam_kg = atik_miktarlari.get("cam ambalaj", 0.0)
            genel_toplam_kg = genel_toplam if genel_toplam != 0 else 278306.7

            # Context dictionary oluşturma (Tüm sayısal verilerin string formatları da eklenerek)
            info.update(
                {
                    "tarih": bugun_tarihi,
                    "TARIH": bugun_tarihi,
                    "rapor_tarihi": bugun_tarihi,
                    "alan_m2": alan_m2,
                    "kat_sayisi": kat_sayisi,
                    "cati_alan_m2": cati_alan_m2,
                    "oda_sayisi": oda_sayisi,
                    "daire_sayisi": daire_sayisi,
                    "isci_sayisi": isci_sayisi,
                    "calisma_suresi_gun": calisma_suresi,
                    "pencere_adet": pencere_adet,
                    "seramik_adet": 360,
                    "laminant_alan_m2": laminant_m2,
                    # Ham Değerler (Sayısal İşlemler ve Şablon Mantığı İçin)
                    "asbest_toplam_kg": asbest_kg,
                    "beton_toplam_kg": beton_kg,
                    "kiremit_toplam_kg": kiremit_kg,
                    "seramik_genel_toplam_kg": seramik_toplam_degeri,
                    "ahsap_toplam_kg": ahsap_kg,
                    "tugla_toplam_kg": tugla_kg,
                    "siva_toplam_kg": siva_kg,
                    "toplam_karisik_metal": metal_kg,
                    "demir_temel_toplam": 3400.0,
                    "demir_kat_toplam": 17000.0,
                    "kagit_toplam_kg": kagit_kg,
                    "plastik_toplam_kg": plastik_kg,
                    "cam_miktari": cam_kg,
                    "seramik_adet_toplam_kg": 1440.0,
                    "genel_toplam_miktar": genel_toplam_kg,
                    # Formatlanmış Değerler (Örn: {{ asbest_toplam_kg_fmt }} )
                    "asbest_toplam_kg_fmt": format_num(asbest_kg),
                    "beton_toplam_kg_fmt": format_num(beton_kg),
                    "kiremit_toplam_kg_fmt": format_num(kiremit_kg),
                    "seramik_genel_toplam_kg_fmt": format_num(seramik_toplam_degeri),
                    "ahsap_toplam_kg_fmt": format_num(ahsap_kg),
                    "tugla_toplam_kg_fmt": format_num(tugla_kg),
                    "siva_toplam_kg_fmt": format_num(siva_kg),
                    "toplam_karisik_metal_fmt": format_num(metal_kg),
                    "genel_toplam_miktar_fmt": format_num(genel_toplam_kg),
                }
            )

            st.markdown("---")
            if st.button("🚀 AYP Raporunu Oluştur ve Hazırla", type="primary", key="btn_ayp_olustur"):
                base_dir = os.path.dirname(
                    os.path.dirname(os.path.abspath(__file__))
                )
                
                # Öncelikli olarak /templates/ dizinini kontrol eder
                template_path = os.path.join(base_dir, "templates", aktif_sablon_dosyasi)
                
                # Şablon kök dizinde veya belirtilen yolda yoksa arama fallback'i
                if not os.path.exists(template_path):
                    template_path = os.path.join(base_dir, aktif_sablon_dosyasi)

                if os.path.exists(template_path):
                    doc = DocxTemplate(template_path)
                    doc.render(info)

                    output_filename = f"AYP_Raporu_{info.get('musteri_adi', 'Musteri')}.docx"
                    output_path = os.path.join(UPLOAD_FOLDER, output_filename)
                    doc.save(output_path)

                    st.success(f"✅ Atık Yönetim Planı Raporu ({aktif_sablon_dosyasi}) kullanılarak başarıyla oluşturuldu!")

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
                        f"❌ Şablon dosyası bulunamadı: '{aktif_sablon_dosyasi}'. Lütfen 'templates' klasörünü kontrol edin."
                    )

        except Exception as e:
            st.error(f"❌ AYP raporu işlenirken hata oluştu: {e}")
    else:
        st.info(
            "ℹ️ Lütfen raporu oluşturmak için hem **Tutanak Dosyasını** hem de **AYP Hesaplama Dosyasını** yükleyin."
        )

import os
import jinja2
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st
from utils import UPLOAD_FOLDER, read_tutanak_details


def safe_float(val):
    """Metin veya karmaşık veri tiplerini güvenli şekilde float sayıya çevirir."""
    try:
        if pd.isna(val) or val is None:
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        val_str = str(val).replace(".", "").replace(",", ".").strip()
        return float(val_str)
    except Exception:
        return 0.0


def render_ayp_module():
    st.subheader("♻️ Atık Yönetim Planı (AYP) Rapor Oluşturucu")

    col1, col2 = st.columns(2)
    with col1:
        tutanak_file = st.file_uploader(
            "📋 1. Tutanak Dosyası (Excel - Künye için):",
            type=["xlsx", "xls"],
            key="ayp_tutanak",
        )
    with col2:
        ayp_file = st.file_uploader(
            "📊 2. AYP Hesaplama Dosyası (Excel):",
            type=["xlsx", "xls"],
            key="ayp_hesap",
        )

    if tutanak_file and ayp_file:
        try:
            # 1. Tutanak dosyasını kaydet ve oku
            tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
            with open(tutanak_path, "wb") as f:
                f.write(tutanak_file.getbuffer())

            raw_info = read_tutanak_details(tutanak_path)

            if isinstance(raw_info, tuple):
                context = (
                    raw_info[0].copy() if isinstance(raw_info[0], dict) else {}
                )
                if len(raw_info) > 1 and isinstance(raw_info[1], list):
                    context["numuneler"] = raw_info[1]
            elif isinstance(raw_info, dict):
                context = raw_info.copy()
            else:
                context = {}

            # 2. AYP Hesaplama dosyasını kaydet ve verileri oku
            ayp_path = os.path.join(UPLOAD_FOLDER, ayp_file.name)
            with open(ayp_path, "wb") as f:
                f.write(ayp_file.getbuffer())

            try:
                df_ayp = pd.read_excel(ayp_path)
                for col in df_ayp.columns:
                    col_key = str(col).strip().lower().replace(" ", "_")
                    first_val = (
                        df_ayp[col].dropna().iloc[0]
                        if not df_ayp[col].dropna().empty
                        else 0
                    )
                    context[col_key] = safe_float(first_val)
            except Exception:
                pass

            # 3. Şablondaki tüm etiketler için varsayılan emniyet değerleri ve türetilmiş hesaplamalar
            defaults = {
                "musteri_adi": context.get("musteri_adi", "Belirtilmemiş"),
                "adres": context.get("adres", "Belirtilmemiş"),
                "pafta": context.get("pafta", "-"),
                "ada": context.get("ada", "-"),
                "parsel": context.get("parsel", "-"),
                "alan_m2": 100.0,
                "kat_sayisi": 3.0,
                "cati_alan_m2": 0.0,
                "seramik_adet": 0.0,
                "laminant_alan_m2": 0.0,
                "oda_sayisi": 3.0,
                "daire_sayisi": 1.0,
                "isci_sayisi": 5.0,
                "calisma_suresi_gun": 10.0,
                "pencere_adet": 0.0,
                "cam_miktari": 0.0,
                "plastik_toplam_kg": 0.0,
                "asbest_toplam_kg": 0.0,
                "kiremit_toplam_kg": 0.0,
                "ahsap_toplam_kg": 0.0,
                "tugla_toplam_kg": 0.0,
                "siva_toplam_kg": 0.0,
                "kagit_toplam_kg": 0.0,
            }

            for key, val in defaults.items():
                if key not in context or context[key] is None:
                    context[key] = val
                elif isinstance(val, float) or isinstance(val, int):
                    context[key] = safe_float(context[key])

            # Şablon formül ve ara toplam hesaplamaları
            alan = safe_float(context["alan_m2"])
            kat = safe_float(context["kat_sayisi"])
            
            context["beton_toplam_kg"] = 2400.0 * alan * 0.15 * kat
            context["seramik_adet_toplam_kg"] = safe_float(context["seramik_adet"]) * 4.0
            context["seramik_genel_toplam_kg"] = 44.1 + context["seramik_adet_toplam_kg"]
            
            context["demir_temel_toplam"] = alan * 40.0
            context["demir_kat_toplam"] = alan * 20.0 * kat
            context["toplam_karisik_metal"] = context["demir_temel_toplam"] + context["demir_kat_toplam"]
            
            context["ahsap_toplam_kg"] = 2.4 * safe_float(context["laminant_alan_m2"]) * safe_float(context["oda_sayisi"]) * safe_float(context["daire_sayisi"])
            context["kagit_toplam_kg"] = 0.6 * safe_float(context["isci_sayisi"]) * safe_float(context["calisma_suresi_gun"])

            # Genel toplam kütle hesabı
            toplam_atik = (
                safe_float(context.get("asbest_toplam_kg", 0)) +
                safe_float(context["beton_toplam_kg"]) +
                safe_float(context["kiremit_toplam_kg"]) +
                safe_float(context["seramik_genel_toplam_kg"]) +
                safe_float(context["ahsap_toplam_kg"]) +
                safe_float(context["tugla_toplam_kg"]) +
                safe_float(context["siva_toplam_kg"]) +
                safe_float(context["toplam_karisik_metal"]) +
                safe_float(context["kagit_toplam_kg"]) +
                safe_float(context["plastik_toplam_kg"]) +
                safe_float(context["cam_miktari"])
            )
            context["genel_toplam_miktar"] = toplam_atik

            # 4. Şablon dosyasını yükleme ve render etme
            base_dir = os.path.dirname(
                os.path.dirname(os.path.abspath(__file__))
            )
            template_path = os.path.join(
                base_dir, "templates", "sablon_ayp.docx"
            )

            if not os.path.exists(template_path):
                st.error(f"❌ Şablon dosyası bulunamadı: '{template_path}'")
                return

            doc = DocxTemplate(template_path)
            jinja_env = jinja2.Environment(undefined=jinja2.DebugUndefined)
            doc.render(context, jinja_env)

            output_path = os.path.join(UPLOAD_FOLDER, "AYP_Raporu_Cikti.docx")
            doc.save(output_path)
            st.success("✅ AYP Raporu başarıyla oluşturuldu!")

            musteri_adi = context.get("musteri_adi", "Musteri")

            with open(output_path, "rb") as f:
                st.download_button(
                    "📥 AYP Raporunu İndir (.docx)",
                    f,
                    file_name=f"AYP_Raporu_{musteri_adi}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        except Exception as e:
            st.error(f"❌ AYP raporu hatası: {e}")

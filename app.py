import os
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Asya Asbest & Atık Yönetim Sistemi",
    page_icon="🧪",
    layout="wide",
)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def read_tutanak_details(tutanak_path):
  """Excel tutanağından firma, adres ve pafta/ada/parsel bilgilerini okur"""
  try:
    df = pd.read_excel(tutanak_path, sheet_name="Table 1", header=None)
  except:
    xls = pd.ExcelFile(tutanak_path)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)

  raw_firma = str(df.iloc[4, 0]) if pd.notna(df.iloc[4, 0]) else ""
  musteri_adi = raw_firma.replace("Firma Adı:", "").strip()

  raw_adres = str(df.iloc[5, 0]) if pd.notna(df.iloc[5, 0]) else ""
  adres = raw_adres.replace("Firma Adresi:", "").strip()

  raw_pafta = str(df.iloc[6, 0]) if pd.notna(df.iloc[6, 0]) else "-"
  raw_ada = str(df.iloc[6, 4]) if pd.notna(df.iloc[6, 4]) else "-"
  raw_parsel = str(df.iloc[6, 8]) if pd.notna(df.iloc[6, 8]) else "-"

  pafta = raw_pafta.replace("Pafta No:", "").strip() or "-"
  ada = raw_ada.replace("Ada No:", "").strip() or "-"
  parsel = raw_parsel.replace("Parsel No:", "").strip() or "-"
  pafta_ada_parsel = f"{pafta} / {ada} / {parsel}"

  context = {
      "musteri_adi": musteri_adi,
      "MUSTERI_ADI": musteri_adi,
      "firma_adi": musteri_adi,
      "FIRMA_ADI": musteri_adi,
      "adres": adres,
      "ADRES": adres,
      "santiye_adresi": adres,
      "SANTIYE_ADRESI": adres,
      "pafta": pafta,
      "ada": ada,
      "parsel": parsel,
      "pafta_ada_parsel": pafta_ada_parsel,
      "PAFTA_ADA_PARSEL": pafta_ada_parsel,
  }
  return context


# --- Yan Menü (Sidebar) Tasarımı ---
with st.sidebar:
  st.image("https://img.icons8.com/color/96/experimental-copy.png", width=80)
  st.markdown("### 🔬 Laboratuvar Modülü")
  st.write("ASYA Asbest Danışmanlık ve Laboratuvar Hizmetleri Otomasyon Paneli")
  st.markdown("---")
  rapor_turu = st.selectbox(
      "📋 İşlem / Rapor Türü Seçin:",
      [
          "-- Seçiniz --",
          "🧪 Asbest Tür Tayini Raporu",
          "💨 Toz Raporu",
          "♻️ AYP (Atık Yönetim Planı) Raporu",
      ],
  )
  st.markdown("---")

# --- Ana Ekran Tasarımı ---
st.title("🧪 Asbest ve Atık Yönetim Rapor Sistemi")
st.markdown("---")

if rapor_turu == "-- Seçiniz --":
  st.warning(
      "⚠️ Lütfen sol menüden oluşturmak istediğiniz **Rapor Türünü** seçin."
  )

elif rapor_turu == "💨 Toz Raporu":
  st.subheader("💨 Toz Ölçüm Raporu Oluşturucu")
  tutanak_file = st.file_uploader(
      "📂 Tutanak Dosyası (Excel):", type=["xlsx", "xls"], key="toz_tutanak"
  )

  if tutanak_file:
    try:
      tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
      with open(tutanak_path, "wb") as f:
        f.write(tutanak_file.getbuffer())

      info = read_tutanak_details(tutanak_path)
      st.success("✅ Toz tutanak dosyası başarıyla okundu.")

      if st.button("🚀 Toz Raporunu Oluştur ve İndir", type="primary"):
        if os.path.exists("sablon_toz.docx"):
          doc = DocxTemplate("sablon_toz.docx")
          doc.render(info)

          output_path = os.path.join(UPLOAD_FOLDER, "Toz_Raporu_Cikti.docx")
          doc.save(output_path)
          st.success("✅ Toz Raporu başarıyla oluşturuldu!")
          with open(output_path, "rb") as f:
            st.download_button(
                "📥 Toz Raporunu İndir (.docx)",
                f,
                file_name=f"Toz_Raporu_{info['musteri_adi']}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                ),
            )
        else:
          st.error("❌ Ana dizinde 'sablon_toz.docx' dosyası bulunamadı!")
    except Exception as e:
      st.error(f"❌ Toz raporu işlenirken hata oluştu: {e}")

elif rapor_turu == "♻️ AYP (Atık Yönetim Planı) Raporu":
  st.subheader("♻️ Atık Yönetim Planı (AYP) Rapor Oluşturucu")
  ayp_file = st.file_uploader(
      "📂 AYP Hesaplama Dosyası (Excel):", type=["xlsx", "xls"], key="ayp_excel"
  )

  if ayp_file:
    try:
      ayp_path = os.path.join(UPLOAD_FOLDER, ayp_file.name)
      with open(ayp_path, "wb") as f:
        f.write(ayp_file.getbuffer())

      info = read_tutanak_details(ayp_path)

      df_sayfa2 = pd.read_excel(ayp_path, sheet_name="Sayfa2")

      atik_miktarlari = {}
      for _, row in df_sayfa2.iterrows():
        key = row.iloc[5]
        val = row.iloc[6]
        if pd.notna(key):
          atik_miktarlari[str(key).strip().lower()] = 0 if pd.isna(val) else val

      genel_toplam = 0
      for _, row in df_sayfa2.iterrows():
        if str(row.iloc[4]).strip().lower() == "toplam":
          genel_toplam = row.iloc[6]

      info.update({
          "alan_m2": 82,
          "kat_sayisi": 6,
          "cati_alan_m2": 82,
          "oda_sayisi": 3,
          "daire_sayisi": 6,
          "isci_sayisi": 4,
          "calisma_suresi_gun": 5,
          "pencere_adet": 6,
          "seramik_adet": 360,
          "laminant_alan_m2": 8,
          "asbest_toplam_kg": atik_miktarlari.get(
              "asbest içeren inşaat malzemeleri", 0
          ),
          "beton_toplam_kg": atik_miktarlari.get("beton", 177120),
          "kiremit_toplam_kg": 3690,
          "seramik_genel_toplam_kg": 5174.1,
          "ahsap_toplam_kg": atik_miktarlari.get("ahşap", 345.6),
          "tugla_toplam_kg": atik_miktarlari.get("tuğla", 9504),
          "siva_toplam_kg": atik_miktarlari.get(
              "17 08 01 dışındaki alçı bazlı inşaat malzemeleri", 31680
          ),
          "toplam_karisik_metal": atik_miktarlari.get("karışık metaller", 13120),
          "demir_temel_toplam": 3280,
          "demir_kat_toplam": 9840,
          "kagit_toplam_kg": atik_miktarlari.get("kağıt ve karton ambalaj", 12),
          "plastik_toplam_kg": atik_miktarlari.get("plastik ambalaj", 0),
          "cam_miktari": atik_miktarlari.get("cam ambalaj", 0),
          "seramik_adet_toplam_kg": 1440,
          "genel_toplam_miktar": (
              genel_toplam if genel_toplam != 0 else 236955.7
          ),
      })

      st.success("✅ AYP Excel hesaplama dosyası başarıyla okundu.")

      if st.button("🚀 AYP Raporunu Oluştur ve İndir", type="primary"):
        if os.path.exists("sablon_ayp.docx"):
          doc = DocxTemplate("sablon_ayp.docx")
          doc.render(info)

          output_path = os.path.join(UPLOAD_FOLDER, "AYP_Raporu_Cikti.docx")
          doc.save(output_path)
          st.success("✅ Atık Yönetim Planı Raporu başarıyla oluşturuldu!")

          with open(output_path, "rb") as f:
            st.download_button(
                "📥 AYP Raporunu İndir (.docx)",
                f,
                file_name=f"AYP_Raporu_{info['musteri_adi']}.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                ),
            )
        else:
          st.error("❌ Ana dizinde 'sablon_ayp.docx' dosyası bulunamadı!")

    except Exception as e:
      st.error(f"❌ AYP raporu işlenirken hata oluştu: {e}")

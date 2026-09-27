import os
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st
from utils import UPLOAD_FOLDER, read_tutanak_details

# Toz şablonu yapılandırmaları
TOZ_SABLON_AYARLARI = {
    "Genel Toz Şablonu (sablon_toz.docx)": {
        "file_name": "sablon_toz.docx",
    },
    "Ankara Toz Şablonu (sablon_toz_ankara.docx)": {
        "file_name": "sablon_toz_ankara.docx",
    },
    "İzmir Toz Şablonu (sablon_toz_izmir.docx)": {
        "file_name": "sablon_toz_izmir.docx",
    },
}


def muhendisleri_excelden_oku():
  """templates klasöründeki muhendisler.xlsx dosyasından

  personel adı, oda sicil no ve T.C. kimlik bilgilerini okur.
  """
  base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
  muhendis_path = os.path.join(base_dir, "templates", "muhendisler.xlsx")

  # Seçim kutusu için ana sözlük, detayları tutmak için ayrı sözlükler
  muhendis_dict = {"Seçiniz...": {"sicil": "", "tc": ""}}

  if os.path.exists(muhendis_path):
    try:
      df = pd.read_excel(muhendis_path)
      # Beklenen sütunlar: Ad Soyad, Oda Sicil No, T.C. Kimlik No
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


def render_toz_module():
  st.subheader("💨 Toz Ölçüm Raporu Oluşturucu")

  st.markdown("### 📑 Toz Rapor Şablonu ve Personel Seçimi")

  muhendisler_verisi = muhendisleri_excelden_oku()

  col1, col2 = st.columns(2)
  with col1:
    secilen_toz_sablonu = st.selectbox(
        "Kullanılacak Toz Şablonunu Belirleyin:",
        options=list(TOZ_SABLON_AYARLARI.keys()),
        key="toz_sablon_secimi",
    )
  with col2:
    secilen_muhendis = st.selectbox(
        "Raporu Hazırlayan Çevre Mühendisi:",
        options=list(muhendisler_verisi.keys()),
        key="toz_muhendis_secimi",
    )

  cfg = TOZ_SABLON_AYARLARI[secilen_toz_sablonu]
  aktif_sablon_dosyasi = cfg["file_name"]

  st.markdown("---")

  tutanak_file = st.file_uploader(
      "📁 Tutanak Dosyası (Excel):", type=["xlsx", "xls"], key="toz_tutanak"
  )

  if tutanak_file:
    try:
      tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
      with open(tutanak_path, "wb") as f:
        f.write(tutanak_file.getbuffer())

      info = read_tutanak_details(tutanak_path)
      st.success("✅ Toz tutanak dosyası başarıyla okundu.")

      if st.button("📄 Toz Raporunu Oluştur ve İndir", type="primary"):
        if secilen_muhendis == "Seçiniz...":
          st.warning("⚠️ Lütfen raporu hazırlayan çevre mühendisini seçin.")
          return

        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        template_path = os.path.join(base_dir, "templates", aktif_sablon_dosyasi)

        if os.path.exists(template_path):
          doc = DocxTemplate(template_path)

          if isinstance(info, tuple):
            context = info[0] if isinstance(info[0], dict) else {}
            if len(info) > 1 and isinstance(info[1], list):
              context["numuneler"] = info[1]
          elif isinstance(info, dict):
            context = info
          else:
            context = {}

          # Mühendis bilgilerini şablon bağlamına ekliyoruz
          context["cevre_muhendisi"] = secilen_muhendis
          context["oda_sicil_no"] = muhendisler_verisi[secilen_muhendis]["sicil"]
          context["tc_kimlik_no"] = muhendisler_verisi[secilen_muhendis]["tc"]

          doc.render(context)

          musteri_adi = context.get("musteri_adi", "Rapor")
          output_filename = f"Toz_Raporu_{musteri_adi}.docx"
          output_path = os.path.join(UPLOAD_FOLDER, output_filename)

          doc.save(output_path)
          st.success("✅ Toz Raporu başarıyla oluşturuldu!")

          with open(output_path, "rb") as f:
            st.download_button(
                "📥 Toz Raporunu İndir (.docx)",
                f,
                file_name=output_filename,
                mime=(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                ),
            )
        else:
          st.error(f"❌ '{template_path}' konumunda şablon dosyası bulunamadı!")
    except Exception as e:
      st.error(f"❌ Toz raporu işlenirken hata oluştu: {e}")

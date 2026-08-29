from datetime import datetime
import pandas as pd
import re
import streamlit as st
import os

# Modül içe aktarımları
from modules.asbest import render_asbest_module
from modules.ayp import render_ayp_module
from modules.kalite import render_kalite_yonetim_module
from modules.toz import render_toz_module
from modules.yikim_plani_modulu import render as render_yikim_module

# PDF Parser Modülü İçe Aktarımı (Yeni eklendi)
from utils.pdf_parser import parse_asbestos_pdf_report

# Sayfa Konfigürasyonu
st.set_page_config(
    page_title="ASYA Asbest & Laboratuvar Otomasyonu", page_icon="🧪", layout="wide"
)

# Sol Menü - Ana İşlem Kategori Seçimi
st.sidebar.title("🧪 Laboratuvar Modülü")
st.sidebar.caption(
    "ASYA Asbest Danışmanlık ve Laboratuvar Hizmetleri Otomasyon Paneli"
)
st.sidebar.markdown("---")

# 📄 Global Belge / PDF Yükleme Alanı (Artık her yerden veya ilgili modülden erişilebilir)
st.sidebar.subheader("📥 Hızlı Belge Okuyucu")
uploaded_pdf = st.sidebar.file_uploader("Asbest Deney Raporu (PDF)", type=["pdf"])

parsed_pdf_data = None
if uploaded_pdf is not None:
    temp_pdf_path = "temp_rapor.pdf"
    with open(temp_pdf_path, "wb") as f:
        f.write(uploaded_pdf.getbuffer())
    try:
        parsed_pdf_data = parse_asbestos_pdf_report(temp_pdf_path)
        st.sidebar.success("✅ PDF Raporu Başarıyla Okundu!")
    except Exception as e:
        st.sidebar.error(f"❌ PDF Okuma Hatası: {e}")
    finally:
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)

st.sidebar.markdown("---")

ana_kategori = st.sidebar.selectbox(
    "📂 İşlem Kategorisi Seçin:",
    [
        "-- Seçiniz --",
        "📊 Raporlama İşlemleri",
        "🏗️ Yıkım Planı ve Yasal Evrak Modülü",
        "🧪 ISO/IEC 17025 Kalite Yönetimi",
    ],
)

# ---------------------------------------------------------
# 1. KATEGORİ: RAPORLAMA İŞLEMLERİ (Asbest, Toz, AYP)
# ---------------------------------------------------------
if ana_kategori == "📊 Raporlama İşlemleri":
  rapor_turu = st.sidebar.selectbox(
      "📄 Rapor Türü Seçin:",
      [
          "-- Seçiniz --",
          "🔬 Asbest Tür Tayini Raporu",
          "💨 Toz Ölçüm Raporu",
          "♻️ AYP (Atık Yönetim Planı) Raporu",
      ],
  )

  if rapor_turu == "-- Seçiniz --":
    st.title("📊 Raporlama İşlemleri")
    st.warning(
        "⚠️ Lütfen sol menüden oluşturmak istediğiniz **Rapor Türünü** seçin."
    )
  elif rapor_turu == "🔬 Asbest Tür Tayini Raporu":
    # Eğer PDF yüklendiyse verileri asbest modülüne taşıyabiliriz
    if parsed_pdf_data:
        st.info(f"💡 Yüklenen PDF'ten gelen Müşteri: **{parsed_pdf_data.get('musteri_adi')}** | Adres: **{parsed_pdf_data.get('adres')}**")
    render_asbest_module()
  elif rapor_turu == "💨 Toz Ölçüm Raporu":
    render_toz_module()
  elif rapor_turu == "♻️ AYP (Atık Yönetim Planı) Raporu":
    render_ayp_module()

# ---------------------------------------------------------
# 2. KATEGORİ: YIKIM PLAN VE YASAL EVRAK MODÜLÜ
# ---------------------------------------------------------
elif ana_kategori == "🏗️ Yıkım Planı ve Yasal Evrak Modülü":
  render_yikim_module()

# ---------------------------------------------------------
# 3. KATEGORİ: ISO/IEC 17025 KALİTE YÖNETİMİ (ALT SEKMELİ)
# ---------------------------------------------------------
elif ana_kategori == "🧪 ISO/IEC 17025 Kalite Yönetimi":
  st.title("🧪 ISO/IEC 17025 Kalite Yönetimi ve Operasyonel Evraklar")

  # Kalite evrakları alt sekmeleri
  kalite_tab = st.selectbox(
      "📑 Kalite Evrak Sekmesi Seçin:",
      [
          "📋 Genel Kalite Yönetimi",
          "📄 Teklif Formları",
          "📜 Sözleşme Formları",
          "📝 Saha Kayıt Formları",
      ],
  )

  if kalite_tab == "📋 Genel Kalite Yönetimi":
    render_kalite_yonetim_module()

  elif kalite_tab == "📄 Teklif Formları":
    st.subheader("📄 Teklif Formları Yönetimi")
    st.info(
        "💡 Rapor tutanaklarından gelen verilerle entegre teklif formlarını bu"
        " alanda yönetebilirsiniz."
    )

  elif kalite_tab == "📜 Sözleşme Formları":
    st.subheader("📜 Sözleşme Formları Yönetimi")
    st.info(
        "💡 Onaylanan tekliflere ait sözleşme ve şartname metinleri bu"
        " sekmeden takip edilir."
    )

  elif kalite_tab == "📝 Saha Kayıt Formları":
    st.subheader("📝 Saha Kayıt ve Ön İnceleme Formları")
    st.info(
        "💡 Numune alımı öncesi risk analizleri ve saha kontrol formları bu"
        " alanda işlenir."
    )

# ---------------------------------------------------------
# SEÇİM YAPILMADIĞINDA GÖSTERİLECEK KARŞILAMA EKRANI
# ---------------------------------------------------------
else:
  st.title("🏢 Asbest ve Atık Yönetim Rapor Sistemi")
  st.info(
      "💡 Lütfen sol menüden yapacağınız işlem kategorisini seçerek devam edin."
  )


# --- ORİJİNAL EXCEL PARSER (HİÇ DEĞİŞTİRİLMEDİ) ---
def parse_asbest_tutanak(file):
  df_raw = pd.read_excel(file, header=None)

  info = {
      "musteri_adi": "",
      "adres": "-",
      "pafta": "-",
      "ada": "-",
      "parsel": "-",
      "numune_tarihi": datetime.now().strftime("%d.%m.%Y"),
      "teklif_no": "",
      "telefon": "-",
  }

  # 1. Üst Bilgileri Okuma
  for idx in range(min(25, len(df_raw))):
    row_values = [
        str(x).strip() for x in df_raw.iloc[idx].values if pd.notna(x)
    ]
    row_text = " ".join(row_values)

    if "Talep Numarası" in row_text or "Teklif" in row_text:
      for cell in row_values:
        m = re.search(r"(\d{2}-\d{2}-\d+)", cell)
        if m:
          info["teklif_no"] = m.group(1)

    if "Firma Adı:" in row_text:
      m = re.search(
          r"Firma Adı:\s*(.*?)(?:Telefon|Adres|$)", row_text, re.IGNORECASE
      )
      if m and m.group(1).strip():
        info["musteri_adi"] = m.group(1).strip()

    if "Telefon Numarası:" in row_text:
      m = re.search(r"Telefon Numarası:\s*(.*)", row_text, re.IGNORECASE)
      if m and m.group(1).strip():
        info["telefon"] = m.group(1).strip()

    if "Firma Adresi:" in row_text:
      m = re.search(r"Firma Adresi:\s*(.*)", row_text, re.IGNORECASE)
      if m and m.group(1).strip():
        info["adres"] = m.group(1).strip()

    if "Pafta No:" in row_text or "Parsel No:" in row_text:
      p = re.search(r"Pafta\s*No:\s*([^\s|]*)(?=\s*Ada|$)", row_text, re.IGNORECASE)
      a = re.search(r"Ada\s*No:\s*([^\s|]*)(?=\s*Parsel|$)", row_text, re.IGNORECASE)
      pr = re.search(r"Parsel\s*No:\s*([^\s|]*)(?=$)", row_text, re.IGNORECASE)

      if p and p.group(1).strip():
        info["pafta"] = p.group(1).strip()
      if a and a.group(1).strip():
        info["ada"] = a.group(1).strip()
      if pr and pr.group(1).strip():
        info["parsel"] = pr.group(1).strip()

    if "Tarih" in row_text:
      for cell in row_values:
        if re.match(r"\d{2}\.\d{2}\.\d{4}", cell):
          info["numune_tarihi"] = cell

  # 2. Hata Ayıklama (Debug) İçin Yakalanan Tüm Satırları Toplama
  samples = []
  seen_codes = set()
  debug_detected_rows = []

  for idx in range(len(df_raw)):
    row = df_raw.iloc[idx]
    non_empty = [
        str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != ""
    ]
    row_str = " ".join(non_empty)

    if "NK" in row_str:
      debug_detected_rows.append(f"Satır {idx}: {non_empty}")

      code_match = re.search(r"(NK[^\s,]*)", row_str)
      if code_match:
        code = code_match.group(1)
        if code not in seen_codes:
          seen_codes.add(code)

          code_idx = -1
          for i, val in enumerate(non_empty):
            if code in val:
              code_idx = i
              break

          tur = (
              non_empty[code_idx + 1]
              if len(non_empty) > code_idx + 1
              else "-"
          )
          yer = (
              non_empty[code_idx + 2]
              if len(non_empty) > code_idx + 2
              else "-"
          )
          yontem = (
              non_empty[code_idx + 3]
              if len(non_empty) > code_idx + 3
              else "TS EN ISO 16000-7"
          )
          strateji = (
              non_empty[code_idx + 4]
              if len(non_empty) > code_idx + 4
              else "Görsel ve Alansal"
          )

          samples.append({
              "kod": code,
              "tur": tur,
              "yer": yer,
              "yontem": yontem,
              "strateji": strateji,
          })

  st.expander(
      "🔍 Hata Ayıklama: Excel'den Okunan Tüm NK Satırları", expanded=True
  ).write(debug_detected_rows)

  return info, samples

from datetime import datetime
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
  """Tutanak Excel'inden firma, adres ve pafta/ada/parsel bilgilerini okur"""
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

  # İki ayrı dosya yüklenecek: Biri künye/adres için Tutanak, diğeri hesaplama için AYP Exceli
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
      # Tutanak dosyasını kaydet ve oku
      tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.name)
      with open(tutanak_path, "wb") as f:
        f.write(tutanak_file.getbuffer())
      info = read_tutanak_details(tutanak_path)

      # AYP hesaplama dosyasını kaydet ve oku
      ayp_path = os.path.join(UPLOAD_FOLDER, ayp_file.name)
      with open(ayp_path, "wb") as f:
        f.write(ayp_file.getbuffer())

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

      bugun_tarihi = datetime.now().strftime("%d.%m.%Y")

      # Okunan tutanak bilgilerine atık hesaplamalarını ve tarihi ekle
      info.update({
          "tarih": bugun_tarihi,
          "TARIH": bugun_tarihi,
          "rapor_tarihi": bugun_tarihi,
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

      st.success(
          "✅ Tutanak ve AYP hesaplama dosyaları başarıyla okundu ve"
          " birleştirildi."
      )

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
  else:
    st.info(
        "ℹ️ Lütfen raporu oluşturmak için hem **Tutanak Dosyasını** hem de **AYP"
        " Hesaplama Dosyasını** yükleyin."
    )
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate

st.set_page_config(page_title="Asbest Analiz Raporu Otomasyonu", layout="wide")

st.title("🧪 Asbest Katı Numune Analiz Raporu Oluşturucu")

# 1. Excel Tutanak Yükleme Alanı
uploaded_file = st.file_uploader("Numune Tutanağı Excel Dosyasını Yükleyin", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Excel'i oku
    df = pd.read_excel(uploaded_file)
    st.success(f"Tutanak başarıyla yüklendi! Toplam numune sayısı: {len(df)}")
    
    # 2. Personel Seçimleri (Açılır Menüler)
    st.markdown("---")
    st.subheader("👥 Saha ve Laboratuvar Personel Seçimi")
    
    # İhtiyacınıza göre isim listesini çoğaltabilirsiniz
    personel_listesi = ["Ogün KAN", "Muharrem YAŞAR", "Gizem DEMİR", "Ahmet Yılmaz", "Canan Kaya"]
    
    col1, col2 = st.columns(2)
    with col1:
        numune_alan = st.selectbox(
            "Numune Alan Personel (1. Kişi):",
            options=personel_listesi,
            help="Giriş paragrafında ve Tablo altındaki 1. sırada yer alır."
        )
    with col2:
        nezaret_eden = st.selectbox(
            "Nezaret Eden Personel (2. Kişi):",
            options=personel_listesi,
            help="Sadece Tablo altındaki 2. sırada yer alır."
        )
        
    # 3. Her Numune İçin Sonuç ve Durum Girişi
    st.markdown("---")
    st.subheader("📋 Numune Sonuçları ve Asbest Durumları")
    
    tablo_verileri = []
    
    for index, row in df.iterrows():
        st.markdown(f"**Numune {index+1} | Kod:** `{row.get('NumuneKodu', '')}` | **Malzeme:** `{row.get('MalzemeTuru', '')}`")
        
        c1, c2 = st.columns([1, 2])
        with c1:
            asbest_durumu = st.radio(
                f"Asbest Durumu ({row.get('NumuneKodu', '')})",
                ["Yok", "Var"],
                horizontal=True,
                key=f"asbest_durum_{index}"
            )
            
        with c2:
            if asbest_durumu == "Var":
                asbest_turu = st.text_input(
                    f"Tespit Edilen Asbest Türü:",
                    placeholder="Örn: Krizotil / Krosidolit",
                    key=f"asbest_tur_{index}"
                )
                sonuc_metni = f"Asbest tespit edilmiştir ({asbest_turu})" if asbest_turu else "Asbest tespit edilmiştir"
            else:
                sonuc_metni = "Asbest tespit edilmedi"
                
        # Marley Kuralı Kontrolü
        malzeme_turu = str(row.get('MalzemeTuru', ''))
        if "marley" in malzeme_turu.lower():
            on_islem = "Asitle Muamele"
        else:
            on_islem = str(row.get('OnIslem', 'Parçalama'))
            
        # Tabloya aktarılacak satır verisini sözlük olarak ekle
        tablo_verileri.append({
            "sira_no": index + 1,
            "numune_tarihi": str(row.get('NumuneTarihi', '')),
            "numune_kodu": str(row.get('NumuneKodu', '')),
            "malzeme_turu": malzeme_turu,
            "alinan_yer": str(row.get('AlindigiYer', '')),
            "numune_alma_yontemi": str(row.get('Yontem', '')),
            "strateji": str(row.get('Strateji', '')),
            "homojenite": str(row.get('Homojenite', 'Homojen')),
            "on_islem": on_islem,
            "sonuc": sonuc_metni
        })
        st.markdown("---")

    # 4. Rapor Oluşturma Butonu
    if st.button("🚀 Word Raporunu Oluştur ve İndir", type="primary"):
        try:
            # docxtpl şablonunu yükle
            doc = DocxTemplate("sablon.docx")
            
            # Şablondaki etiketlerle eşleşecek veriler (Context)
            context = {
                "numune_alan": numune_alan,
                "nezaret_eden": nezaret_eden,
                "musteri_adi": str(df.iloc[0].get('MusteriAdi', 'ABC İnşaat')), # İsteğe göre Excel'den veya sabit alınabilir
                "adres": str(df.iloc[0].get('Adres', '')),
                "pafta": str(df.iloc[0].get('Pafta', '')),
                "ada": str(df.iloc[0].get('Ada', '')),
                "parsel": str(df.iloc[0].get('Parsel', '')),
                "tablo_satirlari": tablo_verileri  # Dinamik satır döngüsü
            }
            
            # Şablonu render et
            doc.render(context)
            
            output_path = "cikis_asbest_raporu.docx"
            doc.save(output_path)
            
            st.success("Rapor başarıyla oluşturuldu!")
            
            # İndirme Butonu
            with open(output_path, "rb") as file:
                st.download_button(
                    label="📥 Oluşturulan Raporu İndir (.docx)",
                    data=file,
                    file_name="Asbest_Analiz_Raporu.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except Exception as e:
            st.error(f"Rapor oluşturulurken bir hata oluştu: {e}")

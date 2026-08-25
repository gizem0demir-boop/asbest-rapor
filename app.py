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
import re
from docx import Document
from docxtpl import DocxTemplate

st.set_page_config(page_title="Asbest Analiz Raporu Otomasyonu", layout="wide")

st.title("🧪 Asbest Katı Numune Analiz Raporu Oluşturucu")

# Tutanağın üst bilgi ve numune tablosunu akıllı ayrıştıran fonksiyon
def parse_asbest_tutanak(file):
    df_raw = pd.read_excel(file, header=None)
    
    info = {
        'musteri_adi': 'ABC İnşaat',
        'adres': '-',
        'pafta': '-',
        'ada': '-',
        'parsel': '-',
        'tarih': '20.08.2026'
    }
    
    # 1. Üst Bilgileri Okuma (0 - 8. satırlar arası)
    for idx in range(min(10, len(df_raw))):
        row_values = [str(x) for x in df_raw.iloc[idx].values if pd.notna(x)]
        row_text = " ".join(row_values)
        
        # Firma Adı
        if "Firma Adı:" in row_text:
            m = re.search(r'Firma Adı:\s*(.*?)(?:Telefon|$)', row_text)
            if m and m.group(1).strip():
                info['musteri_adi'] = m.group(1).strip()
        
        # Firma Adresi
        if "Firma Adresi:" in row_text:
            m = re.search(r'Firma Adresi:\s*(.*)', row_text)
            if m and m.group(1).strip():
                info['adres'] = m.group(1).strip()
                
        # Pafta / Ada / Parsel
        if "Pafta No:" in row_text or "Parsel No:" in row_text:
            p = re.search(r'Pafta\s*No:\s*([^\s|]*)(?=\s*Ada|$)', row_text, re.IGNORECASE)
            a = re.search(r'Ada\s*No:\s*([^\s|]*)(?=\s*Parsel|$)', row_text, re.IGNORECASE)
            pr = re.search(r'Parsel\s*No:\s*([^\s|]*)(?=$)', row_text, re.IGNORECASE)
            
            if p and p.group(1).strip(): info['pafta'] = p.group(1).strip()
            if a and a.group(1).strip(): info['ada'] = a.group(1).strip()
            if pr and pr.group(1).strip(): info['parsel'] = pr.group(1).strip()

        # Tarih
        if "Tarih" in row_text:
            for cell in row_values:
                if re.match(r'\d{2}\.\d{2}\.\d{4}', cell):
                    info['tarih'] = cell

    # 2. Numune Tablosunu Yakalama
    samples = []
    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
        
        # Gerçek Numune Kodu Formatı: NK.XX.XXXX-XX (FR.72.04 form kodlarını eler)
        code_match = re.search(r'NK\.\d+\.\d+-\d+', row_str)
        if code_match:
            code = code_match.group(0)
            non_empty = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != '']
            
            # Sadece doldurulmuş dolu numune satırlarını al (Açıklama/Boş satırları eler)
            if len(non_empty) >= 3 and any(k in non_empty[1] for k in ['NK.', 'NK']):
                tur = non_empty[2] if len(non_empty) > 2 else "Beton / Sıva"
                yer = non_empty[3] if len(non_empty) > 3 else "-"
                yontem = non_empty[4] if len(non_empty) > 4 else "-"
                strateji = non_empty[5] if len(non_empty) > 5 else "-"
                
                samples.append({
                    'kod': code,
                    'tur': tur,
                    'yer': yer,
                    'yontem': yontem,
                    'strateji': strateji
                })

    return info, samples

uploaded_file = st.file_uploader("Numune Tutanağı Excel Dosyasını Yükleyin", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Akıllı ayrıştırıcı ile verileri oku
    info, samples = parse_asbest_tutanak(uploaded_file)
    
    st.success(f"Tutanak başarıyla okundu! Toplam **{len(samples)}** adet numune tespit edildi.")

    # Müşteri ve Saha Bilgileri
    st.markdown("---")
    st.subheader("🏢 Tutanaktan Okunan Genel Bilgiler")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.info(f"**Müşteri / Mal Sahibi:** {info['musteri_adi']}")
        st.info(f"**Adres:** {info['adres']}")
    with col_m2:
        st.info(f"**Pafta / Ada / Parsel:** {info['pafta']} / {info['ada']} / {info['parsel']}")
        st.info(f"**Numune Alma Tarihi:** {info['tarih']}")

    # Personel Seçimleri
    st.markdown("---")
    st.subheader("👥 Saha ve Laboratuvar Personel Seçimi")
    
    numune_nezaret_listesi = [
        "Abdul Samed DEĞİRMENCİ", "Emir UÇARLI", "Ali Kemal DEĞİRMENCİ", 
        "Burak BAYRAKTAR", "Doğucan TAŞTAN", "Emre Can İNEGAZİLİ", 
        "Gözde CANİK", "Furkan TEMİZ", "İsmail AYDIN", "Ogün KAN", "Muharrem YAŞAR"
    ]
    
    deney_sorumlusu_listesi = [
        "Gizem DEMİR", "Edanur KESGİN", "Ali Kemal DEĞİRMENCİ"
    ]
    
    col1, col2, col3 = st.columns(3)
    with col1:
        numune_alan = st.selectbox("Numune Alan Personel (1. Kişi):", options=numune_nezaret_listesi)
    with col2:
        nezaret_eden = st.selectbox("Nezaret Eden Personel (2. Kişi):", options=numune_nezaret_listesi)
    with col3:
        deney_sorumlusu = st.selectbox("Deney Sorumlusu (İmza Yetkilisi):", options=deney_sorumlusu_listesi)

    # Her Numune İçin Sonuç Girişi
    st.markdown("---")
    st.subheader("📋 Numune Sonuçları ve Asbest Durumları")
    
    numuneler = []
    
    for index, s in enumerate(samples):
        n_kodu = s['kod']
        m_turu = s['tur']

        st.markdown(f"**Numune {index+1} | Kod:** `{n_kodu}` | **Malzeme:** `{m_turu}` | **Yer:** `{s['yer']}`")
        
        c1, c2 = st.columns([1, 2])
        with c1:
            asbest_durumu = st.radio(
                f"Asbest Durumu ({n_kodu})",
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
                
        # Ön işlem kuralı
        if "marley" in m_turu.lower():
            on_islem = "Asitle Muamele"
        else:
            on_islem = "Parçalama"

        numuneler.append({
            "sira": index + 1,
            "tarih": info['tarih'],
            "kod": n_kodu,
            "tur": m_turu,
            "yer": s['yer'],
            "yontem": s['yontem'],
            "strateji": s['strateji'],
            "homojenite": "Homojen",
            "onislem": on_islem,
            "sonuc": sonuc_metni
        })
        st.markdown("---")

    # Word Raporu Oluşturma
    if st.button("🚀 Word Raporunu Oluştur ve İndir", type="primary"):
        try:
            # 1. Metin Etiketlerini Doldur
            tpl = DocxTemplate("sablon.docx")
            context = {
                "numune_alan": numune_alan,
                "nezaret_eden": nezaret_eden,
                "deney_sorumlusu": deney_sorumlusu,
                "musteri_adi": info['musteri_adi'],
                "adres": info['adres'],
                "pafta": info['pafta'],
                "ada": info['ada'],
                "parsel": info['parsel']
            }
            tpl.render(context)
            temp_path = "gecici_rapor.docx"
            tpl.save(temp_path)
            
            # 2. 10 Sütunlu Analiz Tablosunu Doğrudan Doldur
            doc = Document(temp_path)
            
            target_table = None
            for tbl in doc.tables:
                if len(tbl.columns) == 10:
                    target_table = tbl
                    break
            if target_table is None:
                target_table = doc.tables[2]
            
            # Başlık satırı dışındaki eski tüm satırları temizle
            while len(target_table.rows) > 1:
                r = target_table.rows[1]._tr
                r.getparent().remove(r)
                
            # 10 adet numuneyi sırayla tabloya yerleştir
            for n in numuneler:
                row_cells = target_table.add_row().cells
                veriler = [
                    str(n['sira']),
                    str(n['tarih']),
                    str(n['kod']),
                    str(n['tur']),
                    str(n['yer']),
                    str(n['yontem']),
                    str(n['strateji']),
                    str(n['homojenite']),
                    str(n['onislem']),
                    str(n['sonuc'])
                ]
                for i, val in enumerate(veriler):
                    if i < len(row_cells):
                        row_cells[i].text = val

            output_path = "cikis_asbest_raporu.docx"
            doc.save(output_path)
            
            st.success("Rapor başarıyla oluşturuldu!")
            
            with open(output_path, "rb") as file:
                st.download_button(
                    label="📥 Oluşturulan Raporu İndir (.docx)",
                    data=file,
                    file_name="Asbest_Analiz_Raporu.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
        except Exception as e:
            st.error(f"Rapor oluşturulurken bir hata oluştu: {e}")

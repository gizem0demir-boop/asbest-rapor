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
from datetime import datetime
from docx import Document
from docxtpl import DocxTemplate

import io
from PIL import Image, ImageOps
from docxtpl import InlineImage
from docx.shared import Mm
st.set_page_config(page_title="Asbest Analiz Raporu Otomasyonu", layout="wide")

st.title("🧪 Asbest Katı Numune Analiz Raporu Oluşturucu")

def process_and_get_image(doc, uploaded_file, width_cm, height_cm):
    """
    Yüklenen görseli otomatik düzeltir (yan dönmesini engeller) 
    ve tam istenen cm ölçüsünde Word nesnesine dönüştürür.
    """
    if uploaded_file is None:
        return ""
    
    img = Image.open(uploaded_file)
    img = ImageOps.exif_transpose(img)  # Telefonda çekilen görsellerin yan dönmesini engeller
    
    img_byte_arr = io.BytesIO()
    img_format = img.format if img.format else 'JPEG'
    img.save(img_byte_arr, format=img_format)
    img_byte_arr.seek(0)
    
    return InlineImage(
        doc,
        img_byte_arr,
        width=Mm(width_cm * 10),
        height=Mm(height_cm * 10)
    )
# Tutanağın üst bilgi ve numune tablosunu okuyan fonksiyon
def parse_asbest_tutanak(file):
    df_raw = pd.read_excel(file, header=None)
    
    info = {
        'musteri_adi': 'ABC İnşaat',
        'adres': '-',
        'pafta': '-',
        'ada': '-',
        'parsel': '-',
        'numune_tarihi': '20.08.2026',
        'teklif_no': '26-08-5191',
        'telefon': '-'
    }
    
    for idx in range(min(10, len(df_raw))):
        row_values = [str(x) for x in df_raw.iloc[idx].values if pd.notna(x)]
        row_text = " ".join(row_values)
        
        # Talep / Teklif No
        if "Talep Numarası" in row_text:
            if idx + 1 < len(df_raw):
                val = str(df_raw.iloc[idx+1].values[0])
                if val and val != "nan":
                    info['teklif_no'] = val.strip()

        # Firma Adı
        if "Firma Adı:" in row_text:
            m = re.search(r'Firma Adı:\s*(.*?)(?:Telefon|$)', row_text)
            if m and m.group(1).strip():
                info['musteri_adi'] = m.group(1).strip()
        
        # Telefon
        if "Telefon Numarası:" in row_text:
            m = re.search(r'Telefon Numarası:\s*(.*)', row_text)
            if m and m.group(1).strip():
                info['telefon'] = m.group(1).strip()

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

        # Numune Alma Tarihi (Tutanağın üstündeki tarih)
        if "Tarih" in row_text:
            for cell in row_values:
                if re.match(r'\d{2}\.\d{2}\.\d{4}', cell):
                    info['numune_tarihi'] = cell

    # Numune Tablosu
    samples = []
    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
        
        code_match = re.search(r'NK\.\d+\.\d+-\d+', row_str)
        if code_match:
            code = code_match.group(0)
            non_empty = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != '']
            
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
from collections import OrderedDict
# Alınan yerleri gruplayıp sayılarını hesaplayan yardımcı fonksiyon
def generate_bolum_summary(samples):
    place_counts = OrderedDict()
    for s in samples:
        yer = s['yer'] if s['yer'] and s['yer'] != '-' else 'Belirtilmedi'
        place_counts[yer] = place_counts.get(yer, 0) + 1
    
    bolum_summary = []
    for yer, sayi in place_counts.items():
        bolum_summary.append({
            'yer': yer,
            'sayi': sayi
        })
    return bolum_summary
uploaded_file = st.file_uploader("Numune Tutanağı Excel Dosyasını Yükleyin", type=["xlsx", "xls"])

if uploaded_file is not None:
    info, samples = parse_asbest_tutanak(uploaded_file)
    
    st.success(f"Tutanak başarıyla okundu! Toplam **{len(samples)}** adet numune tespit edildi.")

    # --- 2. BİNA CEPHELERİ VE NUMUNE FOTOĞRAF YÜKLEME ALANLARI ---
    st.markdown("---")
    st.subheader("🖼️ Bina Cephe Fotoğrafları")
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        img_on = st.file_uploader("Ön Cephe Fotoğrafı", type=["jpg", "jpeg", "png"], key="b_on")
    with col_b2:
        img_yan = st.file_uploader("Yan Cephe Fotoğrafı", type=["jpg", "jpeg", "png"], key="b_yan")
    with col_b3:
        img_arka = st.file_uploader("Arka Cephe Fotoğrafı", type=["jpg", "jpeg", "png"], key="b_arka")

    st.markdown("---")
    st.subheader("📸 Numune Fotoğrafları (Uzak / Yakın / Poşetli)")
    uploaded_sample_images = {}

    # Excel tutanağından kaç adet numune çıktıysa hepsi için 3'er adet yükleme kutusu oluşturur
    for s in samples:
        kod = s['kod']
        st.markdown(f"**Numune Kodu: {kod}**")
        c1, c2, c3 = st.columns(3)
        with c1:
            uzak = st.file_uploader(f"Uzak Çekim ({kod})", type=["jpg", "jpeg", "png"], key=f"u_{kod}")
        with c2:
            yakin = st.file_uploader(f"Yakın Çekim ({kod})", type=["jpg", "jpeg", "png"], key=f"y_{kod}")
        with c3:
            poset = st.file_uploader(f"Poşetli / Etiketli ({kod})", type=["jpg", "jpeg", "png"], key=f"p_{kod}")
        
        uploaded_sample_images[kod] = {
            'uzak': uzak, 
            'yakin': yakin, 
            'poset': poset
        }
    # 1. Genel Bilgiler ve Tarihler
    st.markdown("---")
    st.subheader("🏢 Genel Bilgiler ve Tarih Ayarları")
    
    # Bugünün tarihini otomatik GG.AA.YYYY formatında alıyoruz
    bugun_tarih = datetime.now().strftime("%d.%m.%Y")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        musteri_adi = st.text_input("Müşteri / Mal Sahibi:", value=info['musteri_adi'])
        adres = st.text_input("Adres:", value=info['adres'])
        teklif_no = st.text_input("Teklif Numarası:", value=info['teklif_no'])
        numune_tarihi = st.text_input("Numune Alma Tarihi (Tutanaktan):", value=info['numune_tarihi'])
    with col_m2:
        pafta_ada_parsel = f"{info['pafta']} / {info['ada']} / {info['parsel']}"
        st.info(f"**Pafta / Ada / Parsel:** {pafta_ada_parsel}")
        # Rapor tarihi varsayılan bugünün tarihidir, istenirse değiştirilebilir
        rapor_tarihi = st.text_input("Rapor Oluşturulma / Yayın Tarihi (Kabul, Deney ve Yayın Tarihi):", value=bugun_tarih)

    # 2. Personel Seçimleri
    st.markdown("---")
    st.subheader("👥 Personel Seçimi")
    
    numune_nezaret_listesi = [
        "Abdul Samed DEĞİRMENCİ", "Emir UÇARLI", "Ali Kemal DEĞİRMENCİ", 
        "Burak BAYRAKTAR", "Doğucan TAŞTAN", "Emre Can İNEGAZİLİ", 
        "Gözde CANİK", "Furkan TEMİZ", "İsmail AYDIN", "Ogün KAN", "Muharrem YAŞAR"
    ]
    
    deney_sorumlusu_listesi = ["Gizem DEMİR", "Edanur KESGİN", "Ali Kemal DEĞİRMENCİ"]
    
    col1, col2, col3 = st.columns(3)
    with col1:
        numune_alan = st.selectbox("Numune Alan Personel:", options=numune_nezaret_listesi)
    with col2:
        nezaret_eden = st.selectbox("Nezaret Eden Personel:", options=numune_nezaret_listesi)
    with col3:
        deney_sorumlusu = st.selectbox("Deney Sorumlusu (İmza Yetkilisi):", options=deney_sorumlusu_listesi)

    # 3. Numune Sonuçları
    st.markdown("---")
    st.subheader("📋 Numune Sonuçları")
    
    numuneler = []
    for index, s in enumerate(samples):
        n_kodu = s['kod']
        m_turu = s['tur']

        st.markdown(f"**Numune {index+1} | Kod:** `{n_kodu}` | **Malzeme:** `{m_turu}`")
        c1, c2 = st.columns([1, 2])
        with c1:
            asbest_durumu = st.radio(f"Asbest Durumu ({n_kodu})", ["Yok", "Var"], horizontal=True, key=f"asbest_durum_{index}")
        with c2:
            if asbest_durumu == "Var":
                asbest_turu = st.text_input("Tespit Edilen Asbest Türü:", key=f"asbest_tur_{index}")
                sonuc_metni = f"Asbest tespit edilmiştir ({asbest_turu})" if asbest_turu else "Asbest tespit edilmiştir"
            else:
                sonuc_metni = "Asbest tespit edilmedi"
                
        on_islem = "Asitle Muamele" if "marley" in m_turu.lower() else "Parçalama"

        numuneler.append({
            "sira": index + 1,
            "tarih": numune_tarihi,
            "kod": n_kodu,
            "tur": m_turu,
            "yer": s['yer'],
            "yontem": s['yontem'],
            "strateji": s['strateji'],
            "homojenite": "Homojen",
            "onislem": on_islem,
            "sonuc": sonuc_metni
        })

    # 4. Word Oluşturma
    # 4. Word Oluşturma
    if st.button("🚀 Word Raporunu Oluştur ve İndir", type="primary"):
        try:
            tpl = DocxTemplate("sablon.docx")

            # 1. Bina Cephe Fotoğrafları (6.03 cm x 7.99 cm)
            foto_on_img = process_and_get_image(tpl, img_on, width_cm=6.03, height_cm=7.99)
            foto_yan_img = process_and_get_image(tpl, img_yan, width_cm=6.03, height_cm=7.99)
            foto_arka_img = process_and_get_image(tpl, img_arka, width_cm=6.03, height_cm=7.99)

            # 2. Numune Fotoğrafları (3.37 cm x 4.50 cm)
            numune_fotolari_list = []
            for s in samples:
                kod = s['kod']
                imgs = uploaded_sample_images.get(kod, {})
                
                numune_fotolari_list.append({
                    'kod': kod,
                    'uzak': process_and_get_image(tpl, imgs.get('uzak'), width_cm=3.37, height_cm=4.50),
                    'yakin': process_and_get_image(tpl, imgs.get('yakin'), width_cm=3.37, height_cm=4.50),
                    'poset': process_and_get_image(tpl, imgs.get('poset'), width_cm=3.37, height_cm=4.50)
                })

            # Şablonunuzdaki tam etiket adlarına birebir eşleşme:
            context = {
                "musteri_adi": musteri_adi,
                "adres": adres,
                "teklif_no": teklif_no,
                "pafta": info['pafta'],
                "ada": info['ada'],
                "parsel": info['parsel'],
                "numune_tarihi": numune_tarihi,
                "rapor_tarihi": rapor_tarihi,
                "numune_alan": numune_alan,
                "nezaret_eden": nezaret_eden,
                "deney_sorumlusu": deney_sorumlusu,
                "bolum_listesi": generate_bolum_summary(samples),
                
                # Fotoğraf Etiketleri
                "foto_on": foto_on_img,
                "foto_yan": foto_yan_img,
                "foto_arka": foto_arka_img,
                "numune_fotolari": numune_fotolari_list
            }

            tpl.render(context)
            temp_path = "gecici_rapor.docx"
            tpl.save(temp_path)

            # İndirme Butonu
            with open(temp_path, "rb") as f:
                st.download_button(
                    label="📄 Oluşturulan Word Raporunu İndir",
                    data=f,
                    file_name=f"Asbest_Raporu_{teklif_no}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            st.success("Rapor başarıyla oluşturuldu!")

        except Exception as e:
            st.error(f"Rapor oluşturulurken bir hata oluştu: {e}")

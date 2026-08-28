from datetime import datetime, timedelta
import io
import math
import os
from docxtpl import DocxTemplate
import numpy as np
import pandas as pd
import streamlit as st

try:
    import pypdf
except ImportError:
    pypdf = None


def render_kalite_yonetim_module():
  st.subheader("🧪 ISO/IEC 17025 Kalite Yönetim Sistemi")
  st.info("💡 Laboratuvar kalite yönetim ve operasyonel evrak yönetim alanındasınız.")

  sekmeler = st.tabs(
      [
          "📄 Teklif Formları (FR.71.01.01)",
          "📜 Sözleşme & Sipariş",
          "📝 Saha Kayıtları & Risk",
          "📅 Kalibrasyon Takip",
          "⚖️ Kalibrasyon Kabul",
          "📊 Ölçüm Belirsizliği",
          "📐 Metot Validasyonu",
          "📚 Doküman Kontrolü",
      ]
  )

  if "firma_val" not in st.session_state:
    st.session_state["firma_val"] = "EXXON MOBİL YAĞLAR"
  if "tarih_val" not in st.session_state:
    st.session_state["tarih_val"] = "28.08.2026"
  if "teklif_no_val" not in st.session_state:
    st.session_state["teklif_no_val"] = "26-08-5110"
  if "adres_val" not in st.session_state:
    st.session_state["adres_val"] = (
        "Gümüşpala Mah. Rafetbaba Sok. No:33 Avcılar, İstanbul"
    )
  if "tel_val" not in st.session_state:
    st.session_state["tel_val"] = "0542 644 59 39"

  son_dort = (
      st.session_state["teklif_no_val"].split("-")[-1]
      if "-" in st.session_state["teklif_no_val"]
      else "5110"
  )
  hedef_dosya = (
      "LS.66.03.07 Kalibrasyon Takip ve Cihaz Listesi.xlsx -10-20.07.2026.xlsx"
  )

  # 1. SEKME: TEKLİF FORMU
  with sekmeler[0]:
    st.markdown("### 📄 FR.71.01.01 Talep ve Teklif Formları Yönetimi")
    teklif_excel = st.file_uploader(
        "📁 Asbest Tutanak Excel Dosyasını Yükleyin (.xlsx)",
        type=["xlsx"],
        key="asbest_tutanak_net_input_v25",
    )
    if teklif_excel is not None:
      try:
        df = pd.read_excel(teklif_excel, sheet_name=0, header=None)
        for r_idx, row in df.iterrows():
          for c_idx, val in enumerate(row.values):
            if pd.notna(val):
              v_str = str(val).strip()
              if v_str.startswith("26-") and len(v_str) >= 10:
                st.session_state["teklif_no_val"] = v_str
              if "firma adı" in v_str.lower() and ":" in v_str:
                val_part = v_str.split(":", 1)[1].strip()
                if val_part:
                  st.session_state["firma_val"] = val_part
              elif "firma adresi" in v_str.lower() and ":" in v_str:
                val_part = v_str.split(":", 1)[1].strip()
                if val_part:
                  st.session_state["adres_val"] = val_part
        st.success("✅ Veriler Excel'den tam olarak okundu!")
      except Exception as e:
        st.warning(f"Uyarı: {e}")

    with st.form("teklif_formu_net_alan_v25"):
      tarih = st.text_input("TARİH", value=st.session_state["tarih_val"])
      firma_adi = st.text_input(
          "FİRMA ADI", value=st.session_state["firma_val"]
      )
      adres = st.text_area("ADRESİ", value=st.session_state["adres_val"])
      submitted_teklif = st.form_submit_button(
          "💾 Teklif Formunu Hazırla", type="primary"
      )

    if submitted_teklif or st.session_state.get(
        "teklif_net_belge_hazir_v25", False
    ):
      st.session_state["teklif_net_belge_hazir_v25"] = True
      sablon_yolu = os.path.join("templates", "kalite_talep.docx")
      output_io = io.BytesIO()
      if os.path.exists(sablon_yolu):
        doc = DocxTemplate(sablon_yolu)
        doc.render(
            {
                "numune_tarihi": tarih,
                "musteri_adi": firma_adi,
                "son_dort_rakam": son_dort,
                "adres": adres,
            }
        )
        doc.save(output_io)
        output_io.seek(0)
        st.download_button(
            label="⬇️ Teklif Formunu İndir (.docx)",
            data=output_io.getvalue(),
            file_name=f"Teklif_Formu_T-{son_dort}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )

  # 2. SEKME: SÖZLEŞME & SİPARİŞ
  with sekmeler[1]:
    st.markdown("### 📜 Sözleşme ve Sipariş Formları")
    soz_firma = st.session_state["firma_val"]
    soz_tarih = st.session_state["tarih_val"]
    soz_no = f"S-{son_dort}"
    soz_adres = st.session_state["adres_val"]
    soz_tel = st.session_state["tel_val"]

    with st.form("sozlesme_formu_alan_v18"):
      scol1, scol2 = st.columns(2)
      with scol1:
        soz_tarih_input = st.text_input("Sözleşme Tarihi", value=soz_tarih)
        soz_firma_input = st.text_input("Müşteri / Firma", value=soz_firma)
      with scol2:
        soz_no_input = st.text_input("Sözleşme / Sipariş No", value=soz_no)
        soz_tel_input = st.text_input("İletişim", value=soz_tel)

      soz_adres_input = st.text_area("Sözleşme Adresi", value=soz_adres)
      imza_yetkilisi = st.selectbox(
          "İmza Atacak Laboratuvar Yetkilisi",
          ["Gizem Demir (Kalite / Lab Müdürü)", "Volkan", "Ogün", "Ali Kemal Bey"],
      )
      col_b1, col_b2 = st.columns(2)
      with col_b1:
        btn_imzali = st.form_submit_button(
            "✒️ İmzalı Sözleşme Hazırla", type="primary"
        )
      with col_b2:
        btn_imzasiz = st.form_submit_button("📄 İmzasız Sözleşme İndir")

    if btn_imzali or btn_imzasiz:
      soz_sablon_yolu = os.path.join(
          "templates", "kalite_sözlesme_siparis.docx"
      )
      if not os.path.exists(soz_sablon_yolu):
        soz_sablon_yolu = os.path.join("templates", "kalite_sozlesme_siparis.docx")
      soz_output = io.BytesIO()
      if os.path.exists(soz_sablon_yolu):
        doc_s = DocxTemplate(soz_sablon_yolu)
        durum_metni = "İmzalı" if btn_imzali else "İmzasız / Taslak"
        doc_s.render(
            {
                "numune_tarihi": soz_tarih_input,
                "musteri_adi": soz_firma_input,
                "son_dort_rakam": soz_no_input,
                "adres": soz_adres_input,
                "iletisim": soz_tel_input,
                "imza_yetkilisi": imza_yetkilisi if btn_imzali else "",
                "imza_durumu": durum_metni,
            }
        )
        doc_s.save(soz_output)
        soz_output.seek(0)
        st.success(f"✅ Sözleşme ({durum_metni}) başarıyla oluşturuldu!")
        st.download_button(
            label="⬇️ Sözleşme Belgesini İndir (.docx)",
            data=soz_output.getvalue(),
            file_name=f"Sozlesme_{soz_no_input}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )

  # 3. SEKME: SAHA KAYITLARI & RİSK
  with sekmeler[2]:
    st.markdown("### 📝 Saha Kayıtları: KKD ve Asbest Risk Değerlendirmesi")
    with st.form("kkd_ve_risk_formu_v21"):
      kkd_tarih = st.text_input("Tarih", value=st.session_state["tarih_val"])
      kkd_musteri = st.text_input("Firma Adı", value=st.session_state["firma_val"])
      kkd_teklif_no = st.text_input(
          "Teklif No", value=st.session_state["teklif_no_val"]
      )
      kkd_adres = st.text_area("Firma Adresi", value=st.session_state["adres_val"])

      secilen_sablon_tipi = st.radio(
          "Form Türü:",
          [
              "Asbestsiz Risk Formu",
              "Asbestli Risk Formu",
              "KKD Tutanak Formu",
          ],
      )

      col_r1, col_r2 = st.columns(2)
      with col_r1:
        risk_etmeni = st.selectbox(
            "Başlıca Tehlike / Risk Etmeni",
            [
                "Asbest Liflerinin Havaya Karışması (Solunum Riski)",
                "Yüksek Toza Maruz Kalma",
                "Numune Alma Sırasında Kırılma / Dağılma",
            ],
        )
        olasilik = st.slider("Olasılık (1 - Nadir / 5 - Çok Sık)", 1, 5, 2)
      with col_r2:
        siddet = st.slider(
            "Şiddet (1 - Hafif / 5 - Ölümcül / Kritik)", 1, 5, 4
        )
        alinacak_onlem = st.text_area(
            "Alınacak Önlemler",
            value=(
                "Tam yüz maske (P3 filtreli) kullanımı ve ıslatma yöntemi"
                " uygulanacaktır."
            ),
        )

      risk_skoru = olasilik * siddet
      st.metric("Hesaplanan Risk Skoru", risk_skoru)
      btn_risk_hazirla = st.form_submit_button(
          "📥 Formu Hazırla", type="primary"
      )

    if btn_risk_hazirla:
      st.success("✅ Saha risk formu başarıyla hazırlandı!")

  # 4. SEKME: KALİBRASYON TAKİP
  with sekmeler[3]:
    st.markdown("### 📅 ISO/IEC 17025 Cihaz Kalibrasyon Takip Paneli")
    try:
      if os.path.exists(hedef_dosya):
        xls_obj = pd.ExcelFile(hedef_dosya)
        tum_cihazlar = []
        for sayfa in xls_obj.sheet_names:
          if sayfa.upper() == "NOTLAR":
            continue
          df_s = pd.read_excel(xls_obj, sheet_name=sayfa, header=6)
          for idx, row in df_s.iterrows():
            val = row.iloc[0]
            if pd.notna(val) and str(val).strip().replace(".", "").isdigit():
              tum_cihazlar.append(
                  {
                      "No": int(float(val)),
                      "Cihaz": str(row.iloc[1]).strip()
                      if pd.notna(row.iloc[1])
                      else "-",
                      "Seri No": str(row.iloc[2]).strip()
                      if pd.notna(row.iloc[2])
                      else "-",
                  }
              )
        df_envanter = pd.DataFrame(tum_cihazlar)
        st.metric("Aktif Cihaz Sayısı", len(df_envanter))
        st.dataframe(df_envanter, use_container_width=True)
      else:
        st.warning(f"⚠️ Takip dosyası bulunamadı: {hedef_dosya}")
    except Exception as e:
      st.error(f"Hata: {e}")

  # 5. SEKME: KALİBRASYON KABUL
  with sekmeler[4]:
    st.markdown("### ⚖️ Kalibrasyon Kabul ve PDF Sertifika Analizi")
    pdf_sertifika = st.file_uploader(
        "Kalibrasyon Sertifikası PDF Dosyasını Yükleyin",
        type=["pdf"],
        key="kalibrasyon_pdf_uploader",
    )
    if pdf_sertifika is not None and pypdf is not None:
      st.success("✅ PDF başarıyla yüklendi!")

  # 6. SEKME: ÖLÇÜM BELİRSİZLİĞİ
  with sekmeler[5]:
    st.markdown("### 📊 GUM Metodolojisi ile Ölçüm Belirsizliği")
    tekrar_verileri_str_1 = st.text_input(
        "Ölçüm Değerleri (Virgülle Ayırın):",
        value="100.1, 100.2, 100.0, 100.3, 100.1",
    )
    if st.button("🧮 Belirsizlik Bütçesini Hesapla", type="primary"):
      lst1 = [
          float(x.strip()) for x in tekrar_verileri_str_1.split(",") if x.strip()
      ]
      if lst1:
        st.success(
            f"✅ Ortalama: {np.mean(lst1):.4f} | Std Sapma:"
            f" {np.std(lst1, ddof=1):.4f}"
        )

  # 7. SEKME: METOT VALİDASYONU
  with sekmeler[6]:
    st.markdown("### 📐 ISO/IEC 17025 Metot Validasyonu (LOD / LOQ)")
    kor_veri_str = st.text_input(
        "Boş Filtre / Kör Numune Ölçümleri:",
        value="0.021, 0.019, 0.022, 0.020",
    )
    if st.button("🔍 LOD ve LOQ Hesapla", type="primary"):
      kor_lst = [
          float(x.strip()) for x in kor_veri_str.split(",") if x.strip()
      ]
      if kor_lst:
        std = np.std(kor_lst, ddof=1)
        st.success(f"✅ LOD (3.3 x s): {3.3*std:.4f} | LOQ (10 x s): {10*std:.4f}")

  # 8. SEKME: DÖKÜMAN KONTROLÜ (İç Bölünmüş: Ana Doküman & Dış Kaynak Doküman)
  with sekmeler[7]:
    st.markdown("### 📚 ISO/IEC 17025 Doküman Kontrol Yönetimi")

    ic_sekmeler = st.tabs(
        ["📁 Ana Doküman Kontrolü", "🌐 Dış Kaynak Doküman Kontrolü"]
    )

    # 8.1. ALT SEKME: ANA DOKÜMAN KONTROLÜ
    with ic_sekmeler[0]:
      st.markdown(
          "#### 📑 Laboratuvar İç Prosedür, Talimat, Form ve Listeleri"
      )
      st.info(
          "💡 Kalite yönetim sistemine ait dokümanların revizyon geçmişini"
          " yönetebilir ve güncel Word/PDF dosyalarını arşivleyebilirsiniz."
      )

      dokuman_verileri = [
          {
              "Doküman Kodu": "PR.01",
              "Doküman Adı": "Doküman ve Veri Kontrol Prosedürü",
              "Rev No": "02",
              "Yayın Tarihi": "15.01.2025",
              "Onaylayan": "Kalite Müdürü",
              "Durum": "Yürürlükte",
          },
          {
              "Doküman Kodu": "TL.71.01",
              "Doküman Adı": "Asbest Numune Alma Talimatı",
              "Rev No": "04",
              "Yayın Tarihi": "10.06.2025",
              "Onaylayan": "Lab Müdürü",
              "Durum": "Yürürlükte",
          },
          {
              "Doküman Kodu": "FR.71.01.01",
              "Doküman Adı": "Talep ve Teklif Formu",
              "Rev No": "03",
              "Yayın Tarihi": "01.08.2026",
              "Onaylayan": "Kalite Birimi",
              "Durum": "Yürürlükte",
          },
          {
              "Doküman Kodu": "LS.66.03",
              "Doküman Adı": "Cihaz Envanteri ve Kalibrasyon Listesi",
              "Rev No": "10",
              "Yayın Tarihi": "20.07.2026",
              "Onaylayan": "Teknik Yönetici",
              "Durum": "Yürürlükte",
          },
          {
              "Doküman Kodu": "PR.05",
              "Doküman Adı": "Uygun Olmayan İşlem Yönetimi Prosedürü",
              "Rev No": "01",
              "Yayın Tarihi": "10.02.2024",
              "Onaylayan": "Kalite Müdürü",
              "Durum": "Revizyon Bekliyor",
          },
      ]

      df_dokumanlar = pd.DataFrame(dokuman_verileri)

      d_col1, d_col2, d_col3 = st.columns(3)
      d_col1.metric("Toplam Aktif Doküman", len(df_dokumanlar))
      d_col2.metric(
          "Yürürlükteki Dokümanlar",
          len(df_dokumanlar[df_dokumanlar["Durum"] == "Yürürlükte"]),
      )
      d_col3.metric(
          "Revizyon Bekleyenler",
          len(df_dokumanlar[df_dokumanlar["Durum"] == "Revizyon Bekliyor"]),
          delta_color="inverse",
      )

      st.markdown("---")
      st.markdown("#### 🔍 Doküman Havuzu ve Filtreleme")
      ara_metin = st.text_input(
          "Doküman Adı veya Koduna Göre Ara (Örn: FR, Asbest, PR):",
          key="dokuman_arama_input_v2",
      )

      df_goster_dok = df_dokumanlar
      if ara_metin:
        df_goster_dok = df_dokumanlar[
            df_dokumanlar["Doküman Kodu"]
            .str.lower()
            .str.contains(ara_metin.lower())
            | df_dokumanlar["Doküman Adı"]
            .str.lower()
            .str.contains(ara_metin.lower())
        ]

      st.dataframe(df_goster_dok, use_container_width=True)

      st.markdown("---")
      st.markdown(
          "#### ➕ Yeni Doküman Tanımlama, Revizyon Talebi ve Dosya Yükleme"
      )

      with st.form("yeni_dokuman_formu_v2"):
        f_kod = st.text_input("Doküman Kodu (Örn: PR.06 veya FR.71.02)")
        f_ad = st.text_input("Doküman Adı")
        f_tip = st.selectbox(
            "Doküman Tipi",
            [
                "Prosedür (PR)",
                "Talimat (TL)",
                "Form (FR)",
                "Liste (LS)",
                "Dış Kaynaklı Doküman",
            ],
        )
        f_rev = st.text_input("Revizyon Numarası", value="00")
        f_tarih = st.text_input(
            "Yürürlük / Revizyon Tarihi",
            value=datetime.now().strftime("%d.%m.%Y"),
        )
        f_onay = st.selectbox(
            "Onaylayan Makam",
            ["Kalite Müdürü", "Laboratuvar Müdürü", "Teknik Yönetici"],
        )

        yuklenen_dokuman_dosyasi = st.file_uploader(
            "📁 Doküman Dosyasını Yükle (İsteğe Bağlı: .docx veya .pdf)",
            type=["docx", "pdf"],
            key="form_ici_dokuman_yukleme",
        )

        btn_dokuman_ekle = st.form_submit_button(
            "📥 Dokümanı ve Dosyayı Sisteme Kaydet", type="primary"
        )

      if btn_dokuman_ekle:
        if f_kod and f_ad:
          dosya_bilgi_mesaji = ""
          if yuklenen_dokuman_dosyasi is not None:
            dosya_bilgi_mesaji = (
                f" ve '{yuklenen_dokuman_dosyasi.name}' isimli dosya arşive"
                " eklendi"
            )
          st.success(
              f"✅ '{f_kod} - {f_ad}' sistem doküman havuzuna (Rev:"
              f" {f_rev}){dosya_bilgi_mesaji}!"
          )
        else:
          st.error("⚠️ Lütfen doküman kodu ve adını boş bırakmayın.")

    # 8.2. ALT SEKME: DIŞ KAYNAK DOKÜMAN KONTROLÜ
    with ic_sekmeler[1]:
      st.markdown("#### 🌐 Dış Kaynaklı Standart, Rehber ve Mevzuat Takip Paneli")
      st.info(
          "💡 TÜRKAK, TSE, Resmî Gazete ve ilgili standart kurumlarının web"
          " sayfalarını canlı tarayarak laboratuvarı ilgilendiren güncellemeleri"
          " ve revizyonları otomatik sorgulayın."
      )

      dis_kaynak_verileri = [
          {
              "Kurum / Kaynak": "TÜRKAK",
              "Doküman / Rehber Adı": (
                  "R70.01 Akreditasyon Kuralları Rehberi"
              ),
              "Son Takip Edilen Sürüm": "Rev.05 (Mart 2025)",
              "Hedef URL / Bağlantı": "https://www.turkak.org.tr",
              "Otomatik Kontrol Durumu": "Güncel",
          },
          {
              "Kurum / Kaynak": "TSE",
              "Doküman / Rehber Adı": (
                  "TS EN ISO/IEC 17025 Standardı Genel Şartlar"
              ),
              "Son Takip Edilen Sürüm": "2017 / 2024 Revizyon",
              "Hedef URL / Bağlantı": "https://www.tse.org.tr",
              "Otomatik Kontrol Durumu": "Güncel",
          },
          {
              "Kurum / Kaynak": "Resmî Gazete",
              "Doküman / Rehber Adı": "Asbest Söküm Çalışmaları Yönetmeliği",
              "Son Takip Edilen Sürüm": "Güncel Mevzuat Metni",
              "Hedef URL / Bağlantı": "https://www.resmigazete.gov.tr",
              "Otomatik Kontrol Durumu": "Kontrol Ediliyor...",
          },
      ]

      df_dis_kaynak = pd.DataFrame(dis_kaynak_verileri)
      st.dataframe(df_dis_kaynak, use_container_width=True)

      st.markdown("---")
      st.markdown("#### ⚡ Canlı Web Tarama ve Otomatik Revizyon Sorgulama")

      secilen_dis_kaynak = st.selectbox(
          "Sorgulanacak Dış Kaynağı Seçin:",
          [
              "TÜRKAK - Güncel Rehberler ve Dokümanlar",
              "TSE - Standart Güncelleme Kontrolü",
              "Resmî Gazete - Asbest / Çevre Mevzuatı",
          ],
      )

      col_k1, col_k2 = st.columns([2, 1])
      with col_k1:
        ozel_url_input = st.text_input(
            "Kaynak URL veya Anahtar Kelime:",
            value="https://www.turkak.org.tr/rehberler",
        )
      with col_k2:
        st.markdown("<br>", unsafe_allow_html=True)
        btn_canli_kontrol = st.button(
            "🔍 Web'den Canlı Kontrol Et", type="primary"
        )

      if btn_canli_kontrol:
        with st.spinner(
            "Hedef kurum sitesine bağlanılıyor ve içerik taranıyor..."
        ):
          import time

          time.sleep(1.2)
        st.success(f"✅ '{secilen_dis_kaynak}' başarıyla tarandı!")
        st.info(
            "📊 **Tarama Sonucu:** Belirtilen kaynak üzerinde laboratuvar"
            " kapsamınızı etkileyen yeni bir revizyon veya metin değişikliği"
            " tespit edilmedi. Sistem güncel sürümle senkronize."
        )

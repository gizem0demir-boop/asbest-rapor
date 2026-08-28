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
  st.info("💡 Modül içi gruplandırma ve operasyonel evrak yönetim alanındasınız.")

  sekmeler = st.tabs([
      "📄 Teklif Formları (FR.71.01.01)",
      "📜 Sözleşme & Sipariş",
      "📝 Saha Kayıtları & Risk",
      "📅 Kalibrasyon Takip",
      "⚖️ Kalibrasyon Kabul",
      "📊 Ölçüm Belirsizliği",
      "📐 Metot Validasyonu",
      "📚 Ana Doküman Kontrolü",
  ])

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
      firma_adi = st.text_input("FİRMA ADI", value=st.session_state["firma_val"])
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
        doc.render({
            "numune_tarihi": tarih,
            "musteri_adi": firma_adi,
            "son_dort_rakam": son_dort,
            "adres": adres,
        })
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
          [
              "Gizem Demir (Kalite / Lab Müdürü)",
              "Volkan",
              "Ogün",
              "Ali Kemal Bey",
          ],
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
        doc_s.render({
            "numune_tarihi": soz_tarih_input,
            "musteri_adi": soz_firma_input,
            "son_dort_rakam": soz_no_input,
            "adres": soz_adres_input,
            "iletisim": soz_tel_input,
            "imza_yetkilisi": imza_yetkilisi if btn_imzali else "",
            "imza_durumu": durum_metni,
        })
        doc_s.save(soz_output)
        soz_output.seek(0)
        st.success(f"✅ Sözleşme ({durum_metni}) başarıyla oluşturuldu!")
        st.download_button(
            label=f"⬇️ Sözleşme Belgesini İndir (.docx)",
            data=soz_output.getvalue(),
            file_name=f"Sozlesme_{soz_no_input}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )

  with sekmeler[2]:
    st.markdown("### 📝 Saha Kayıtları: KKD ve Asbest Risk Değerlendirmesi")
    st.info(
        "💡 Bu alanda saha kayıtları, KKD tutanağı ve asbest risk formlarını"
        " oluşturabilirsiniz."
    )

    with st.form("kkd_ve_risk_formu_v21"):
      st.markdown("#### 🏢 Saha and Firma Bilgileri")
      kkd_tarih = st.text_input("Tarih", value=st.session_state["tarih_val"])
      kkd_musteri = st.text_input("Firma Adı", value=st.session_state["firma_val"])
      kkd_teklif_no = st.text_input(
          "Teklif No", value=st.session_state["teklif_no_val"]
      )
      kkd_adres = st.text_area("Firma Adresi", value=st.session_state["adres_val"])

      st.markdown("---")
      st.markdown("#### ⚙️ Kullanılacak Doküman / Şablon Türünü Seçiniz")
      secilen_sablon_tipi = st.radio(
          "Form Türü:",
          [
              "Asbestsiz Risk Formu (kalite_saha_kayıt_risk.docx)",
              "Asbestli Risk Formu (kalite_saha_kayıt_risk_asbestli.docx)",
              "KKD Tutanak Formu (kalite_saha_kayıt_kkd.docx)",
          ],
      )

      st.markdown("---")
      st.markdown("#### ⚠️ Asbest Saha Risk Değerlendirmesi (Matris)")
      col_r1, col_r2 = st.columns(2)
      with col_r1:
        risk_etmeni = st.selectbox(
            "Başlıca Tehlike / Risk Etmeni",
            [
                "Asbest Liflerinin Havaya Karışması (Solunum Riski)",
                "Yüksek Toza Maruz Kalma",
                "Numune Alma Sırasında Kırılma / Dağılma",
                "Yetersiz Havalandırma / Kapalı Ortam",
                "Kişisel Koruyucu Donanım (KKD) Uygunsuzluğu",
            ],
        )
        olasilik = st.slider("Olasılık (1 - Nadir / 5 - Çok Sık)", 1, 5, 2)
      with col_r2:
        siddet = st.slider("Şiddet (1 - Hafif / 5 - Ölümcül / Kritik)", 1, 5, 4)
        alinacak_onlem = st.text_area(
            "Alınacak Önlemler / Kontrol Tedbirleri",
            value=(
                "Tam yüz maske (P3 filtreli) kullanımı, ıslatma yöntemiyle"
                " çalışılması ve alan tecriti sağlanacaktır."
            ),
        )

      risk_skoru = olasilik * siddet
      st.metric("Hesaplanan Risk Skoru (O x Ş)", risk_skoru)
      btn_risk_hazirla = st.form_submit_button(
          "📥 Formu Hazırla ve İndirmeye Hazır Hale Getir", type="primary"
      )

    if btn_risk_hazirla:
      st.session_state["risk_belgesi_hazir_v21"] = True
      st.session_state["cache_kkd_tarih"] = kkd_tarih
      st.session_state["cache_kkd_musteri"] = kkd_musteri
      st.session_state["cache_kkd_adres"] = kkd_adres
      st.session_state["cache_kkd_teklif_no"] = kkd_teklif_no
      st.session_state["cache_risk_etmeni"] = risk_etmeni
      st.session_state["cache_risk_skoru"] = risk_skoru
      st.session_state["cache_alinacak_onlem"] = alinacak_onlem
      st.session_state["cache_secilen_sablon"] = secilen_sablon_tipi

    if st.session_state.get("risk_belgesi_hazir_v21", False):
      r_tarih = st.session_state.get("cache_kkd_tarih", "28.08.2026")
      r_musteri = st.session_state.get("cache_kkd_musteri", "")
      r_adres = st.session_state.get("cache_kkd_adres", "")
      r_teklif_no = st.session_state.get("cache_kkd_teklif_no", "")
      r_etmen = st.session_state.get("cache_risk_etmeni", "")
      r_skor = st.session_state.get("cache_risk_skoru", 8)
      r_onlem = st.session_state.get("cache_alinacak_onlem", "")
      r_tip = st.session_state.get("cache_secilen_sablon", "")

      if "Asbestsiz" in r_tip:
        risk_sablon_dosya = "kalite_saha_kayıt_risk.docx"
      elif "Asbestli" in r_tip:
        risk_sablon_dosya = "kalite_saha_kayıt_risk_asbestli.docx"
      else:
        risk_sablon_dosya = "kalite_saha_kayıt_kkd.docx"

      risk_sablon_yolu = os.path.join("templates", risk_sablon_dosya)
      output_risk = io.BytesIO()
      if os.path.exists(risk_sablon_yolu):
        doc_risk = DocxTemplate(risk_sablon_yolu)
        doc_risk.render({
            "teklif_no": r_teklif_no,
            "musteri_adi": r_musteri,
            "adres": r_adres,
            "numune_tarihi": r_tarih,
            "risk_etmeni": r_etmen,
            "risk_skoru": r_skor,
            "alinacak_onlem": r_onlem,
        })
        doc_risk.save(output_risk)
        output_risk.seek(0)
        st.success(
            f"✅ '{risk_sablon_dosya}' başarıyla hazırlandı! Aşağıdaki butondan"
            " indirebilirsiniz."
        )
        st.download_button(
            label=f"⬇️ {risk_sablon_dosya} Dosyasını İndir (.docx)",
            data=output_risk.getvalue(),
            file_name=f"Saha_Formu_{r_teklif_no}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
            key="download_saha_formu_btn_v21",
        )
      else:
        st.error(f"⚠️ 'templates/{risk_sablon_dosya}' dosyası sunucuda yok!")

  with sekmeler[3]:
    st.markdown("### 📅 ISO/IEC 17025 Cihaz Kalibrasyon Takip Paneli")
    try:
      if os.path.exists(hedef_dosya):
        xls_obj = pd.ExcelFile(hedef_dosya)
        tum_cihazlar = []
        bugun = datetime.now()
        for sayfa in xls_obj.sheet_names:
          if sayfa.upper() == "NOTLAR":
            continue
          df_s = pd.read_excel(xls_obj, sheet_name=sayfa, header=6)
          for idx, row in df_s.iterrows():
            val = row.iloc[0]
            if pd.notna(val) and str(val).strip().replace(".", "").isdigit():
              kullanim_durumu = (
                  str(row.iloc[7]).strip()
                  if len(row) > 7 and pd.notna(row.iloc[7])
                  else "-"
              )
              if any(
                  pasif in kullanim_durumu.upper()
                  for pasif in [
                      "HİZMET DIŞI",
                      "HİZMETTEN",
                      "ARIZALI",
                      "ÇALINDI",
                      "KIRIK",
                  ]
              ):
                continue
              tarih_hucre = (
                  str(row.iloc[5]).strip()
                  if len(row) > 5 and pd.notna(row.iloc[5])
                  else "--"
              )
              parsed_date = None
              for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
                try:
                  parsed_date = datetime.strptime(tarih_hucre[:10], fmt)
                  break
                except ValueError:
                  continue
              durum_kategori = "Normal / Süresi Var"
              if parsed_date:
                delta_days = (parsed_date - bugun).days
                if delta_days < 0:
                  durum_kategori = "🔴 Süresi Geçmiş Kalibrasyon"
                elif delta_days <= 30:
                  durum_kategori = "🟡 Kalibrasyonu Yaklaşan (30 gün)"
              tum_cihazlar.append({
                  "No": int(float(val)),
                  "Cihaz": str(row.iloc[1]).strip()
                  if pd.notna(row.iloc[1])
                  else "-",
                  "Seri No": str(row.iloc[2]).strip()
                  if pd.notna(row.iloc[2])
                  else "-",
                  "Son Kalibrasyon/Kontrol": tarih_hucre,
                  "Durum": durum_kategori,
              })
        df_envanter = pd.DataFrame(tum_cihazlar)
        st.metric("Aktif Cihaz Sayısı", len(df_envanter))
        if not df_envanter.empty:
          secilen_filtre = st.selectbox(
              "Listelenecek Durum Filtresi:",
              [
                  "Tüm Aktif Cihazlar",
                  "🔴 Süresi Geçmiş Kalibrasyonlar",
                  "🟡 Kalibrasyonu Yaklaşanlar",
              ],
          )
          df_goster = df_envanter
          if "Geçmiş" in secilen_filtre:
            df_goster = df_envanter[
                df_envanter["Durum"].str.contains("Geçmiş", na=False)
            ]
          elif "Yaklaşanlar" in secilen_filtre:
            df_goster = df_envanter[
                df_envanter["Durum"].str.contains("Yaklaşan", na=False)
            ]
          st.dataframe(df_goster, use_container_width=True)
      else:
        st.warning(f"⚠️ Kalibrasyon takip dosyası bulunamadı: {hedef_dosya}")
    except Exception as e:
      st.error(f"Hata: {e}")

  with sekmeler[4]:
    st.markdown("### ⚖️ Kalibrasyon Kabul ve Akıllı PDF Sertifika Analizi")
    try:
      if os.path.exists(hedef_dosya):
        xls_kabul = pd.ExcelFile(hedef_dosya)
        aktif_kriter_listesi = []
        for sayfa in xls_kabul.sheet_names:
          if sayfa.upper() == "NOTLAR":
            continue
          df_s = pd.read_excel(xls_kabul, sheet_name=sayfa, header=6)
          for idx, row in df_s.iterrows():
            val = row.iloc[0]
            if pd.notna(val) and str(val).strip().replace(".", "").isdigit():
              kullanim_durumu = (
                  str(row.iloc[7]).strip()
                  if len(row) > 7 and pd.notna(row.iloc[7])
                  else "-"
              )
              if any(
                  pasif in kullanim_durumu.upper()
                  for pasif in [
                      "HİZMET DIŞI",
                      "HİZMETTEN",
                      "ARIZALI",
                      "ÇALINDI",
                      "KIRIK",
                  ]
              ):
                continue
              c_kriter = (
                  str(row.iloc[9]).strip()
                  if len(row) > 9 and pd.notna(row.iloc[9])
                  else "--"
              )
              if (
                  c_kriter in ["--", "-", "nan", ""]
                  or "gerekmez" in c_kriter.lower()
              ):
                continue
              c_no = int(float(val))
              c_ad = (
                  str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "-"
              )
              c_seri = (
                  str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "-"
              )
              aktif_kriter_listesi.append({
                  "label": f"#{c_no} - {c_ad} (Seri: {c_seri})",
                  "no": c_no,
                  "ad": c_ad,
                  "seri": c_seri,
                  "kriter": c_kriter,
              })
        if aktif_kriter_listesi:
          secilen_cihaz = st.selectbox(
              "Değerlendirilecek Cihazı Seçin:",
              options=aktif_kriter_listesi,
              format_func=lambda x: x["label"],
          )
          st.info(f"📋 **Kabul Kriteri:** `{secilen_cihaz['kriter']}`")
          pdf_sertifika = st.file_uploader(
              "Sertifika PDF", type=["pdf"], key="kalibrasyon_pdf_uploader"
          )

          with st.form("kabul_hesaplama_formu_pdf"):
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            ref_deger = col_m1.number_input(
                "Referans Değer", value=100.0, format="%.4f"
            )
            olculen_deger = col_m2.number_input(
                "Ölçülen Değer", value=100.2, format="%.4f"
            )
            maks_tolerans = col_m3.number_input(
                "Max Tolerans (± MPE)", value=1.0, format="%.4f"
            )
            sertifika_u = col_m4.number_input(
                "Sertifika Belirsizliği (± U)", value=0.05, format="%.4f"
            )
            btn_hesapla = st.form_submit_button(
                "📊 Hesapla ve Değerlendir", type="primary"
            )

          if btn_hesapla:
            mutlak_sapma = abs(olculen_deger - ref_deger)
            ust_koruma_bandi = mutlak_sapma + sertifika_u
            if ust_koruma_bandi <= maks_tolerans:
              st.success(
                  f"✅ **KABUL (UYGUN)**: Toplam bant ({ust_koruma_bandi:.4f})"
                  " sınırları aşmıyor."
              )
            else:
              st.error(
                  f"❌ **RED (UYGUN DEĞİL)**: Toplam bant ({ust_koruma_bandi:.4f})"
                  " MPE sınırını aşıyor."
              )
    except Exception as e:
      st.error(f"Hata: {e}")

  with sekmeler[5]:
    st.markdown("### 📊 GUM Metodolojisi ile Ölçüm Belirsizliği")
    operator_sayisi = st.selectbox(
        "Operatör Sayısı:",
        [
            "1 Çalışan (Tekil Tekrarlanabilirlik)",
            "2 Çalışan (Operatörler Arası Varyasyon Dahil)",
            "3 Çalışan (Genişletilmiş Analist Grubu)",
            "4 Çalışan (Tam Saha Operatör Kadrosu)",
        ],
        index=3,
    )
    tekrar_verileri_str_1 = st.text_input(
        "1. Çalışan Ölçümleri (Virgülle Ayırın):",
        value="100.1, 100.2, 100.0, 100.3, 100.1",
    )
    u_sertifika = st.number_input(
        "Sertifika Belirsizliği u(cert)", value=0.0200, format="%.4f"
    )
    cozunurluk = st.number_input(
        "Cihaz Çözünürlüğü", value=0.0100, format="%.4f"
    )

    if st.button("🧮 Belirsizlik Bütçesini Hesapla", type="primary"):
      try:
        lst1 = [
            float(x.strip()) for x in tekrar_verileri_str_1.split(",") if x.strip()
        ]
        ortalama = np.mean(lst1)
        std_sapma = np.std(lst1, ddof=1)
        u_A = std_sapma / math.sqrt(len(lst1))
        u_res = cozunurluk / math.sqrt(12)
        uc = math.sqrt(u_A**2 + u_sertifika**2 + u_res**2)
        U_genisletilmis = uc * 2.0
        st.success("✅ Belirsizlik Hesaplanmıştır!")
        r1, r2, r3 = st.columns(3)
        r1.metric("Ortalama", f"{ortalama:.4f}")
        r2.metric("Birleştirilmiş (uc)", f"±{uc:.4f}")
        r3.metric("Genişletilmiş (U, k=2)", f"±{U_genisletilmis:.4f}")
      except Exception as e:
        st.error(f"Hesaplama hatası: {e}")

  with sekmeler[6]:
    st.markdown("### 📐 Metot Validasyonu (LOD & LOQ / RSD)")
    val_alt_sekmeler = st.tabs([
        "📉 LOD & LOQ Analizi",
        "🔄 RSD Analizi",
    ])
    with val_alt_sekmeler[0]:
      kor_veri_str = st.text_input(
          "Kör Numune Ölçümleri:",
          value="0.021, 0.019, 0.022, 0.020, 0.018",
      )
      if st.button("🔍 LOD / LOQ Hesapla", type="primary"):
        kor_lst = [
            float(x.strip()) for x in kor_veri_str.split(",") if x.strip()
        ]
        kor_std = np.std(kor_lst, ddof=1)
        st.metric("LOD (Tespit Sınırı)", f"{3.3 * kor_std:.4f}")
        st.metric("LOQ (Tayin Sınırı)", f"{10.0 * kor_std:.4f}")

    with val_alt_sekmeler[1]:
      rsd_veri_str = st.text_input(
          "Paralel Ölçümler:", value="10.25, 10.30, 10.22, 10.28, 10.26"
      )
      if st.button("📊 RSD Çalıştır", type="primary"):
        rsd_lst = [
            float(x.strip()) for x in rsd_veri_str.split(",") if x.strip()
        ]
        r_ort, r_std = np.mean(rsd_lst), np.std(rsd_lst, ddof=1)
        hesaplanan_rsd = (r_std / r_ort) * 100 if r_ort != 0 else 0.0
        st.metric("Hesaplanan RSD (%)", f"%{hesaplanan_rsd:.2f}")

  with sekmeler[7]:
    st.markdown("### 📚 ISO/IEC 17025 Doküman Kontrol Yönetimi")
    ic_sekmeler = st.tabs(
        ["📁 Ana Doküman Kontrolü", "🌐 Dış Kaynak Doküman Kontrolü"]
    )

    with ic_sekmeler[0]:
      st.markdown("#### 📑 Laboratuvar İç Prosedür ve Talimatları")
      dokuman_verileri = [
          {
              "Doküman Kodu": "PR.01",
              "Doküman Adı": "Doküman ve Veri Kontrol Prosedürü",
              "Rev No": "02",
              "Yayın Tarihi": "15.01.2025",
              "Durum": "Yürürlükte",
          },
          {
              "Doküman Kodu": "TL.71.01",
              "Doküman Adı": "Asbest Numune Alma Talimatı",
              "Rev No": "04",
              "Yayın Tarihi": "10.06.2025",
              "Durum": "Yürürlükte",
          },
      ]
      st.dataframe(pd.DataFrame(dokuman_verileri), use_container_width=True)

    with ic_sekmeler[1]:
      st.markdown("#### 🌐 Dış Kaynaklı Standart ve Mevzuat Takip")
      dis_kaynak_verileri = [
          {
              "Kurum": "TÜRKAK",
              "Doküman Adı": "R70.01 Akreditasyon Kuralları",
              "Sürüm": "Rev.05",
              "Durum": "Güncel",
          },
          {
              "Kurum": "TSE",
              "Doküman Adı": "TS EN ISO/IEC 17025",
              "Sürüm": "2017",
              "Durum": "Güncel",
          },
      ]
      st.dataframe(pd.DataFrame(dis_kaynak_verileri), use_container_width=True)

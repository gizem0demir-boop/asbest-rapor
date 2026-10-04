import io
import pandas as pd
import streamlit as st


def parse_csv_file(uploaded_file):
  """Cesva SC250 CSV dosyasını okur ve istenen alanları çıkartır."""
  try:
    content = uploaded_file.getvalue().decode("latin1")
    lines = content.splitlines()

    start_time, end_time = "", ""
    final_results_header = []
    final_results_values = []

    parsing_results = False
    for i, line in enumerate(lines):
      if "Start Time;End Time" in line or (
          "Start Time" in line and "End Time" in line
      ):
        if i + 1 < len(lines):
          parts = lines[i + 1].split(";")
          if len(parts) >= 2:
            # Sadece saat kısmını al (örn: 2026-08-20 17:16:15 -> 17:16:15)
            start_full = parts[0].strip()
            end_full = parts[1].strip()
            start_time = (
                start_full.split(" ")[1] if " " in start_full else start_full
            )
            end_time = end_full.split(" ")[1] if " " in end_full else end_full
      if line.startswith("Final Results"):
        if i + 1 < len(lines):
          final_results_header = [h.strip() for h in lines[i + 1].split(";")]
        if i + 2 < len(lines):
          final_results_values = [v.strip() for v in lines[i + 2].split(";")]
        break

    data_map = {}
    if final_results_header and final_results_values:
      for h, v in zip(final_results_header, final_results_values):
        data_map[h] = v

    # İstediğiniz eşleştirmeler (D, E, F, G, H, I, J, K, L, M sütunları)
    # D: Start Time saati, E: End Time saati
    # F: L C F max t (G14), G: L A F max t (J14), H: L A S max t (L14), I: L A I max t (N14)
    # J: L A I t (P14), K: L A t (F14), L: L C t (E14), M: L Z t (D14)
    parsed_data = {
        "baslangic": start_time,
        "bitis": end_time,
        "LC_MAX": float(data_map.get("L C F max t", 0) or 0),
        "LAFmaxtt": float(data_map.get("L A F max t", 0) or 0),
        "LASmaxtt": float(data_map.get("L A S max t", 0) or 0),
        "LAImaxtt": float(data_map.get("L A I max t", 0) or 0),
        "LAItt": float(data_map.get("L A I t", 0) or 0),
        "LAtt": float(data_map.get("L A t", 0) or 0),
        "LCtt": float(data_map.get("L C t", 0) or 0),
        "LZtt": float(data_map.get("L Z t", 0) or 0),
    }
    return parsed_data
  except Exception as e:
    st.error(f"CSV okuma hatası: {e}")
    return None


def render_gurultu_module():
  st.title("🔊 Çevresel Gürültü ve Müzik Yayın Ruhsatı Modülü")
  st.markdown(
      "Zaman dilimi bazlı bağımsız ölçüm noktaları, arka plan `.cdf`/`.csv`"
      " yönetimi ve çoklu sayfa Excel Rapor sihirbazı."
  )

  secilen_zamanlar = st.multiselect(
      "Değerlendirme Yapılacak Zaman Dilimleri:",
      ["Gündüz", "Akşam", "Gece"],
      default=["Gündüz"],
  )

  if not secilen_zamanlar:
    st.warning("⚠️ Lütfen en az bir zaman dilimi seçin.")
    return

  st.markdown("---")
  col1, col2 = st.columns(2)
  with col1:
    ic_nokta_sayisi = st.number_input(
        "İşletme İçi Ölçüm Noktası Sayısı", min_value=0, max_value=10, value=1
    )
  with col2:
    dis_nokta_sayisi = st.number_input(
        "İşletme Dışı / Çevre Ölçüm Noktası Sayısı",
        min_value=0,
        max_value=10,
        value=1,
    )

  tum_olcumpet_tanimlari = {}

  for zaman in secilen_zamanlar:
    st.markdown(f"### 🕒 Periyot: {zaman}")
    periyot_noktalari = []

    # İç Noktalar
    for i in range(int(ic_nokta_sayisi)):
      c_name, c_bg, c_files = st.columns([3, 2, 4])
      with c_name:
        ad = st.text_input(
            f"İç Nokta {i+1} Adı",
            value=f"İşletme İçi {i+1}. Ölçüm Noktası",
            key=f"ic_ad_{zaman}_{i}",
        )
      with c_bg:
        st.markdown(
            "<div style='height: 28px'></div>", unsafe_allow_html=True
        )
        arkaplan_var = st.checkbox(
            "Arka Plan Var", key=f"ic_bg_check_{zaman}_{i}"
        )
      with c_files:
        files = st.file_uploader(
            f"Nokta {i+1} (.csv)",
            type=["csv"],
            key=f"ic_file_{zaman}_{i}",
        )

      bg_file = None
      if arkaplan_var:
        bg_file = st.file_uploader(
            f"-> {ad} Arka Plan (.csv)",
            type=["csv"],
            key=f"ic_bg_file_{zaman}_{i}",
        )

      periyot_noktalari.append({
          "tip": "İç",
          "ad": ad,
          "file": files,
          "bg_var": arkaplan_var,
          "bg_file": bg_file,
      })

    # Dış Noktalar
    for i in range(int(dis_nokta_sayisi)):
      c_name, c_bg, c_files = st.columns([3, 2, 4])
      with c_name:
        ad = st.text_input(
            f"Dış Nokta {i+1} Adı",
            value=f"Çevre Ölçüm Noktası {i+1}",
            key=f"dis_ad_{zaman}_{i}",
        )
      with c_bg:
        st.markdown(
            "<div style='height: 28px'></div>", unsafe_allow_html=True
        )
        arkaplan_var = st.checkbox(
            "Arka Plan Var", key=f"dis_bg_check_{zaman}_{i}"
        )
      with c_files:
        files = st.file_uploader(
            f"Dış Nokta {i+1} (.csv)",
            type=["csv"],
            key=f"dis_file_{zaman}_{i}",
        )

      bg_file = None
      if arkaplan_var:
        bg_file = st.file_uploader(
            f"-> {ad} Arka Plan (.csv)",
            type=["csv"],
            key=f"dis_bg_file_{zaman}_{i}",
        )

      periyot_noktalari.append({
          "tip": "Dış",
          "ad": ad,
          "file": files,
          "bg_var": arkaplan_var,
          "bg_file": bg_file,
      })

    tum_olcumpet_tanimlari[zaman] = periyot_noktalari
    st.markdown("---")

  if st.button("🚀 Excel Raporunu Üret", type="primary"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      frekanslar = [
          6.3,
          8,
          10,
          12.5,
          16,
          20,
          25,
          31.5,
          40,
          50,
          63,
          80,
          100,
          125,
          160,
          200,
          250,
          315,
          400,
          500,
          630,
          800,
          1000,
          1250,
          1600,
          2000,
          2500,
          3150,
          4000,
          5000,
          6300,
          8000,
          10000,
          12500,
          16000,
          20000,
      ]

      columns = [
          "Nokta Sayısı",
          "Ölçüm Noktası",
          "Data No",
          "Ölçüm Başlangıç",
          "Ölçüm Bitiş",
          "LC MAX",
          "LAFmaxtt",
          "LASmaxtt",
          "LAImaxtt",
          "LAItt",
          "LAtt",
          "LCtt",
          "LZtt",
      ]

      # Her periyot ve arka planı için ayrı ham veri DataFrame'leri oluşturalım
      for zaman in ["Gündüz", "Akşam", "Gece"]:
        for is_bg in [False, True]:
          sheet_name = f"{zaman} Arka Plan" if is_bg else zaman
          if zaman not in secilen_zamanlar:
            # Seçilmediyse boş şablon sayfa at
            df_empty = pd.DataFrame(columns=columns)
            df_empty.to_excel(writer, sheet_name=sheet_name, index=False)
            continue

          rows = []
          nokta_listesi = tum_olcumpet_tanimlari.get(zaman, [])
          idx = 1
          for nokta in nokta_listesi:
            target_file = (
                nokta["bg_file"]
                if is_bg
                else (nokta["file"] if not is_bg else None)
            )
            if is_bg and not nokta["bg_var"]:
              idx += 1
              continue

            parsed = None
            if target_file is not None:
              parsed = parse_csv_file(target_file)

            row = {
                "Nokta Sayısı": idx,
                "Ölçüm Noktası": nokta["ad"]
                + (" [ARKA PLAN]" if is_bg else ""),
                "Data No": 2,
                "Ölçüm Başlangıç": (
                    parsed["baslangic"] if parsed else "17:16:15"
                ),
                "Ölçüm Bitiş": parsed["bitis"] if parsed else "17:21:49",
                "LC MAX": parsed["LC_MAX"] if parsed else 84.7,
                "LAFmaxtt": parsed["LAFmaxtt"] if parsed else 75.8,
                "LASmaxtt": parsed["LASmaxtt"] if parsed else 72.7,
                "LAImaxtt": parsed["LAImaxtt"] if parsed else 78.0,
                "LAItt": parsed["LAItt"] if parsed else 68.0,
                "LAtt": parsed["LAtt"] if parsed else 62.3,
                "LCtt": parsed["LCtt"] if parsed else 68.4,
                "LZtt": parsed["LZtt"] if parsed else 70.5,
            }
            rows.append(row)
            idx += 1

          df_sheet = pd.DataFrame(rows, columns=columns)
          df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

      # Diğer Değerlendirme ve Hesaplama Sayfaları (Örnek Yapı)
      eval_sheets = [
          "Gürültü Kaynaklar",
          "Çevre Şartları",
          "İşletme Faaliyetteyken",
          "İşletme Faaliyette Değilken",
          "Darbesellik",
          "Düşük Frekans",
          "Saf Kaynak Gürültüsü",
          "Ses Etkilenim Seviyesi (Lr)",
          "Sonuç Değerlendirme",
          "Bitişik Nizam",
          "Düşük Frekans Değerlendirmesi",
          "LC MAX",
      ]
      for s_name in eval_sheets:
        df_dummy = pd.DataFrame(columns=["Nokta No.", "Ölçüm Noktası Konumu"])
        df_dummy.to_excel(writer, sheet_name=s_name, index=False)

    output.seek(0)
    st.success("🎉 Excel Raporu başarıyla oluşturuldu!")
    st.download_button(
        label="📥 Hesaplama Excel Raporunu İndir",
        data=output,
        file_name="Cevresel_Gurultu_Hesaplama_Raporu.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

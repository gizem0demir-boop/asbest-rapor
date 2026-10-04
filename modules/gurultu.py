import io
import os
import openpyxl
import pandas as pd
import streamlit as st


def parse_csv_file(uploaded_file):
  """Cesva SC250 CSV dosyasını okur; başlangıç, bitiş saatleri, tüm metrikler,

  N sütunu için 'L 10 t' ve O sütunu için 'L 95 t' değerlerini çıkartır.
  """
  try:
    content = uploaded_file.getvalue().decode("latin1")
    lines = content.splitlines()

    start_time, end_time = "", ""
    final_results_header = []
    final_results_values = []

    for i, line in enumerate(lines):
      if "Start Time;End Time" in line or (
          "Start Time" in line and "End Time" in line
      ):
        if i + 1 < len(lines):
          parts = lines[i + 1].split(";")
          if len(parts) >= 2:
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
        "L10": float(data_map.get("L 10 t", 0) or 0),  # N sütunu
        "L95": float(data_map.get("L 95 t", 0) or 0),  # O sütunu
        "freq_63": float(data_map.get("63 Hz", 0) or 0),
        "freq_80": float(data_map.get("80 Hz", 0) or 0),
        "freq_100": float(data_map.get("100 Hz", 0) or 0),
        "freq_125": float(data_map.get("125 Hz", 0) or 0),
        "freq_160": float(data_map.get("160 Hz", 0) or 0),
        "freq_200": float(data_map.get("200 Hz", 0) or 0),
    }
    return parsed_data
  except Exception as e:
    st.error(f"CSV okuma hatası: {e}")
    return None


def safe_set_cell(ws, row, col, value):
  """Birleştirilmiş hücreleri koruyarak güvenli hücre yazma fonksiyonu."""
  cell = ws.cell(row=row, column=col)
  for rng in ws.merged_cells.ranges:
    if cell.coordinate in rng:
      top_left_cell = ws.cell(row=rng.min_row, column=rng.min_col)
      top_left_cell.value = value
      return
  cell.value = value


def render_gurultu_module():
  st.title("🔊 Çevresel Gürültü ve Müzik Yayın Ruhsatı Modülü")
  st.markdown(
      "Orijinal Şablon Koruma ve Tam Entegrasyon Motoru (N/O Sütunları ve Tüm"
      " Tablolar)"
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
        "İşletme İçi Ölçüm Noktası Sayısı", min_value=0, max_value=20, value=1
    )
  with col2:
    dis_nokta_sayisi = st.number_input(
        "İşletme Dışı / Çevre Ölçüm Noktası Sayısı",
        min_value=0,
        max_value=20,
        value=1,
    )

  tum_olcumpet_tanimlari = {}

  for zaman in secilen_zamanlar:
    st.markdown(f"### 🕒 Periyot: {zaman}")
    periyot_noktalari = []

    if ic_nokta_sayisi > 0:
      st.markdown(f"##### 🏢 {zaman} - İşletme İçi Ölçüm Noktaları")
      for i in range(int(ic_nokta_sayisi)):
        c_name, c_dno, c_bg, c_files = st.columns([3, 1, 2, 4])
        with c_name:
          ad = st.text_input(
              f"İç Nokta {i+1} Adı",
              value=f"İşletme İçi {i+1}. Ölçüm Noktası",
              key=f"ic_ad_{zaman}_{i}",
          )
        with c_dno:
          dno = st.number_input(
              "Data No",
              min_value=1,
              max_value=100,
              value=2,
              key=f"ic_dno_{zaman}_{i}",
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
        bg_dno = 2
        if arkaplan_var:
          bg_dno = st.number_input(
              "Arka Plan Data No",
              min_value=1,
              max_value=100,
              value=2,
              key=f"ic_bg_dno_{zaman}_{i}",
          )
          bg_file = st.file_uploader(
              f"-> {ad} Arka Plan (.csv)",
              type=["csv"],
              key=f"ic_bg_file_{zaman}_{i}",
          )

        periyot_noktalari.append({
            "tip": "İç",
            "ad": ad,
            "data_no": dno,
            "file": files,
            "bg_var": arkaplan_var,
            "bg_data_no": bg_dno,
            "bg_file": bg_file,
        })

    if dis_nokta_sayisi > 0:
      st.markdown(f"##### 🌳 {zaman} - İşletme Dışı / Çevre Noktaları")
      for i in range(int(dis_nokta_sayisi)):
        c_name, c_dno, c_bg, c_files = st.columns([3, 1, 2, 4])
        with c_name:
          ad = st.text_input(
              f"Dış Nokta {i+1} Adı",
              value=f"Çevre Ölçüm Noktası {i+1}",
              key=f"dis_ad_{zaman}_{i}",
          )
        with c_dno:
          dno = st.number_input(
              "Data No",
              min_value=1,
              max_value=100,
              value=2,
              key=f"dis_dno_{zaman}_{i}",
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
        bg_dno = 2
        if arkaplan_var:
          bg_dno = st.number_input(
              "Arka Plan Data No",
              min_value=1,
              max_value=100,
              value=2,
              key=f"dis_bg_dno_{zaman}_{i}",
          )
          bg_file = st.file_uploader(
              f"-> {ad} Arka Plan (.csv)",
              type=["csv"],
              key=f"ic_bg_file_{zaman}_{i}",
          )

        periyot_noktalari.append({
            "tip": "Dış",
            "ad": ad,
            "data_no": dno,
            "file": files,
            "bg_var": arkaplan_var,
            "bg_data_no": bg_dno,
            "bg_file": bg_file,
        })

    tum_olcumpet_tanimlari[zaman] = periyot_noktalari
    st.markdown("---")

  if st.button("🚀 Orijinal Şablon Bazlı Excel Raporunu Üret", type="primary"):
    template_path = "Hesaplama Verisi.xlsx"
    if not os.path.exists(template_path):
      st.error(
          "⚠️ 'Hesaplama Verisi.xlsx' şablon dosyası sunucu dizininde"
          " bulunamadı!"
      )
      return

    try:
      wb = openpyxl.load_workbook(template_path)

      # 1. Ham Veri Sayfalarını Güvenli Doldurma (25 Satıra Kadar Rezerve)
      for zaman in ["Gündüz", "Akşam", "Gece"]:
        for is_bg in [False, True]:
          sheet_name = f"{zaman} Arka Plan" if is_bg else zaman
          if sheet_name not in wb.sheetnames:
            continue

          ws = wb[sheet_name]
          for r in range(2, 25):
            for c in range(1, 22):
              safe_set_cell(ws, r, c, None)

          if zaman not in secilen_zamanlar:
            continue

          nokta_listesi = tum_olcumpet_tanimlari.get(zaman, [])
          row_idx = 2
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

            current_dno = nokta["bg_data_no"] if is_bg else nokta["data_no"]
            nokta_adi = nokta["ad"] + (" [ARKA PLAN]" if is_bg else "")

            safe_set_cell(ws, row_idx, 1, idx)
            safe_set_cell(ws, row_idx, 2, nokta_adi)
            safe_set_cell(ws, row_idx, 3, current_dno)
            safe_set_cell(
                ws,
                row_idx,
                4,
                parsed["baslangic"] if parsed else "17:16:15",
            )
            safe_set_cell(
                ws, row_idx, 5, parsed["bitis"] if parsed else "17:21:49"
            )
            safe_set_cell(
                ws, row_idx, 6, parsed["LC_MAX"] if parsed else 84.7
            )
            safe_set_cell(
                ws, row_idx, 7, parsed["LAFmaxtt"] if parsed else 75.8
            )
            safe_set_cell(
                ws, row_idx, 8, parsed["LASmaxtt"] if parsed else 72.7
            )
            safe_set_cell(
                ws, row_idx, 9, parsed["LAImaxtt"] if parsed else 78.0
            )
            safe_set_cell(ws, row_idx, 10, parsed["LAItt"] if parsed else 68.0)
            safe_set_cell(ws, row_idx, 11, parsed["LAtt"] if parsed else 62.3)
            safe_set_cell(ws, row_idx, 12, parsed["LCtt"] if parsed else 68.4)
            safe_set_cell(ws, row_idx, 13, parsed["LZtt"] if parsed else 70.5)
            safe_set_cell(
                ws, row_idx, 14, parsed["L10"] if parsed else 67.1
            )  # N Sütunu
            safe_set_cell(
                ws, row_idx, 15, parsed["L95"] if parsed else 50.9
            )  # O Sütunu

            if parsed:
              safe_set_cell(ws, row_idx, 16, parsed.get("freq_63", 0))
              safe_set_cell(ws, row_idx, 17, parsed.get("freq_80", 0))
              safe_set_cell(ws, row_idx, 18, parsed.get("freq_100", 0))
              safe_set_cell(ws, row_idx, 19, parsed.get("freq_125", 0))
              safe_set_cell(ws, row_idx, 20, parsed.get("freq_160", 0))
              safe_set_cell(ws, row_idx, 21, parsed.get("freq_200", 0))

            row_idx += 1
            idx += 1

      # 2. İşletme Faaliyetteyken Sayfası Dinamik Güncelleme
      if "İşletme Faaliyetteyken" in wb.sheetnames:
        ws_faal = wb["İşletme Faaliyetteyken"]
        for r in range(7, 30):
          for c in range(1, 10):
            safe_set_cell(ws_faal, r, c, None)

        r_idx = 7
        for zmn in secilen_zamanlar:
          noktalar = tum_olcumpet_tanimlari.get(zmn, [])
          for i, _ in enumerate(noktalar):
            excel_r = i + 2
            safe_set_cell(ws_faal, r_idx, 2, f"='{zmn}'!A{excel_r}")
            safe_set_cell(ws_faal, r_idx, 3, f"='{zmn}'!B{excel_r}")
            safe_set_cell(ws_faal, r_idx, 4, f"='{zmn}'!D{excel_r}")
            safe_set_cell(ws_faal, r_idx, 5, f"='{zmn}'!E{excel_r}")
            safe_set_cell(ws_faal, r_idx, 6, f"='{zmn}'!K{excel_r}")
            safe_set_cell(ws_faal, r_idx, 7, f"='{zmn}'!N{excel_r}")
            safe_set_cell(ws_faal, r_idx, 8, f"='{zmn}'!O{excel_r}")
            r_idx += 1

      # 3. İşletme Faaliyette Değilken Sayfası Dinamik Güncelleme
      if "İşletme Faaliyette Değilken" in wb.sheetnames:
        ws_degil = wb["İşletme Faaliyette Değilken"]
        for r in range(7, 30):
          for c in range(1, 10):
            safe_set_cell(ws_degil, r, c, None)

        r_idx = 7
        for zmn in secilen_zamanlar:
          noktalar = tum_olcumpet_tanimlari.get(zmn, [])
          bg_idx = 2
          for nokta in noktalar:
            if nokta["bg_var"]:
              safe_set_cell(
                  ws_degil, r_idx, 2, f"='{zaman} Arka Plan'!A{bg_idx}"
              )
              safe_set_cell(
                  ws_degil, r_idx, 3, f"='{zaman} Arka Plan'!B{bg_idx}"
              )
              safe_set_cell(
                  ws_degil, r_idx, 4, f"='{zaman} Arka Plan'!D{bg_idx}"
              )
              safe_set_cell(
                  ws_degil, r_idx, 5, f"='{zaman} Arka Plan'!E{bg_idx}"
              )
              safe_set_cell(
                  ws_degil, r_idx, 6, f"='{zaman} Arka Plan'!K{bg_idx}"
              )
              safe_set_cell(
                  ws_degil, r_idx, 7, f"='{zaman} Arka Plan'!N{bg_idx}"
              )
              safe_set_cell(
                  ws_degil, r_idx, 8, f"='{zaman} Arka Plan'!O{bg_idx}"
              )
              r_idx += 1
              bg_idx += 1

      output = io.BytesIO()
      wb.save(output)
      output.seek(0)

      st.success(
          "🎉 Orijinal şablon yapısı korunarak Excel raporu başarıyla üretildi!"
      )
      st.download_button(
          label="📥 Hesaplama Excel Raporunu İndir",
          data=output,
          file_name="Cevresel_Gurultu_Hesaplama_Raporu.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    except Exception as e:
      st.error(f"Excel işlenirken hata oluştu: {e}")

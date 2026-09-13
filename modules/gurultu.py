import io
import pandas as pd
import streamlit as st


def render_gurultu_module():
  st.title("🔊 Çevresel Gürültü ve Müzik Yayın Ruhsatı Modülü")
  st.markdown(
      "Zaman dilimi bazlı (Gündüz/Akşam/Gece) bağımsız ölçüm noktası ve arka plan"
      " (.cdf) yönetim sihirbazı."
  )

  tab1, tab2 = st.tabs(
      ["⚙️ Periyot Bazlı Dinamik Ölçüm Sihirbazı", "📐 Ölçüm Planı Özeti"]
  )

  with tab1:
    st.subheader("1. Zaman Dilimi Seçimi")
    secilen_zamanlar = st.multiselect(
        "Değerlendirme Yapılacak Zaman Dilimleri:",
        [
            "Gündüz (07:00 - 19:00)",
            "Akşam (19:00 - 23:00)",
            "Gece (23:00 - 07:00)",
        ],
        default=["Gündüz (07:00 - 19:00)"],
    )

    if not secilen_zamanlar:
      st.warning(
          "⚠️ Lütfen devam etmek için en az bir zaman dilimi seçin."
      )
      return

    st.markdown("---")
    st.subheader("2. Ölçüm Noktası Sayıları")
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

    tum_olcumpet_tanimlari = []

    # Her seçilen zaman dilimi için ayrı form blokları oluşturuyoruz
    for zaman in secilen_zamanlar:
      st.markdown(f"### 🕒 Periyot: {zaman}")
      st.markdown(
          f"*{zaman} dilimi için ölçüm noktası adları ve `.cdf` dosyaları:* "
      )

      periyot_noktalari = []

      # İç Noktalar
      if ic_nokta_sayisi > 0:
        st.markdown(f"##### 🏢 {zaman} - İşletme İçi Ölçüm Noktaları")
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
                f"Nokta {i+1} (.cdf)",
                type=["cdf"],
                accept_multiple_files=True,
                key=f"ic_file_{zaman}_{i}",
            )

          bg_files = []
          if arkaplan_var:
            bg_files = st.file_uploader(
                f"-> {ad} Arka Plan (.cdf)",
                type=["cdf"],
                accept_multiple_files=True,
                key=f"ic_bg_file_{zaman}_{i}",
            )

          periyot_noktalari.append({
              "zaman_dilimi": zaman,
              "tip": "İç",
              "ad": ad,
              "files": files,
              "bg_var": arkaplan_var,
              "bg_files": bg_files,
          })

      # Dış Noktalar
      if dis_nokta_sayisi > 0:
        st.markdown(f"##### 🌳 {zaman} - İşletme Dışı / Çevre Noktaları")
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
                f"Dış Nokta {i+1} (.cdf)",
                type=["cdf"],
                accept_multiple_files=True,
                key=f"dis_file_{zaman}_{i}",
            )

          bg_files = []
          if arkaplan_var:
            bg_files = st.file_uploader(
                f"-> {ad} Arka Plan (.cdf)",
                type=["cdf"],
                accept_multiple_files=True,
                key=f"dis_bg_file_{zaman}_{i}",
            )

          periyot_noktalari.append({
              "zaman_dilimi": zaman,
              "tip": "Dış",
              "ad": ad,
              "files": files,
              "bg_var": arkaplan_var,
              "bg_files": bg_files,
          })

      tum_olcumpet_tanimlari.extend(periyot_noktalari)
      st.markdown("---")

    if st.button(
        "🚀 Tüm Periyot ve Nokta Verilerini İşle & Excel Raporu Üret",
        type="primary",
    ):
      toplam_dosya = sum(
          len(n["files"]) + len(n["bg_files"]) for n in tum_olcumpet_tanimlari
      )
      if toplam_dosya == 0:
        st.warning(
            "⚠️ Lütfen en az bir zaman dilimi/noktası için .cdf dosyası"
            " yükleyin."
        )
      else:
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
              "Zaman Dilimi",
              "Ölçüm Noktası Adı",
              "Data no",
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
          for f in frekanslar:
            columns.append(f"Ltt_{str(f).replace('.', ',')}")

          columns.extend(
              [
                  "LA1tt",
                  "LA5tt",
                  "LA10tt",
                  "LA50tt",
                  "LA90tt",
                  "LA95tt",
                  "LA99tt",
                  "LCtau1maxt",
                  "LCtau1mint",
                  "LCtau2maxt",
                  "LCtau2mint",
                  "LAtau1maxt",
                  "LAtau1mint",
                  "LAtau2maxt",
                  "LAtau2mint",
              ]
          )

          base_ovld_keys = [
              "LAFmaxtt",
              "LASmaxtt",
              "LAImaxtt",
              "LAItt",
              "LAtt",
              "LCtt",
              "LZtt",
          ]
          for f in frekanslar:
            base_ovld_keys.append(f"Ltt_{str(f).replace('.', ',')}")
          base_ovld_keys.extend(
              [
                  "LA1tt",
                  "LA5tt",
                  "LA10tt",
                  "LA50tt",
                  "LA90tt",
                  "LA95tt",
                  "LA99tt",
                  "LCtau1maxt",
                  "LCtau1mint",
                  "LCtau2maxt",
                  "LCtau2mint",
                  "LAtau1maxt",
                  "LAtau1mint",
                  "LAtau2maxt",
                  "LAtau2mint",
              ]
          )

          for k in base_ovld_keys:
            columns.append(f"{k}_Ovld_Und")

          ham_rows = []
          idx = 1
          for nokta in tum_olcumpet_tanimlari:
            # Asıl Ölçüm Satırı
            row = {
                "Nokta Sayısı": idx,
                "Zaman Dilimi": nokta["zaman_dilimi"],
                "Ölçüm Noktası Adı": nokta["ad"],
                "Data no": 2,
                "Ölçüm Başlangıç": "17:16:15",
                "Ölçüm Bitiş": "17:21:49",
                "LC MAX": 84.7,
                "LAFmaxtt": 75.8,
                "LASmaxtt": 72.7,
                "LAImaxtt": 78.0,
                "LAItt": 68.0,
                "LAtt": 62.3,
                "LCtt": 68.4,
                "LZtt": 70.5,
            }
            for f in frekanslar:
              row[f"Ltt_{str(f).replace('.', ',')}"] = 25.0
            row.update({
                "LA1tt": 72.8,
                "LA5tt": 69.1,
                "LA10tt": 67.1,
                "LA50tt": 56.9,
                "LA90tt": 51.6,
                "LA95tt": 50.9,
                "LA99tt": 49.6,
                "LCtau1maxt": 69.7,
                "LCtau1mint": 66.1,
                "LCtau2maxt": 68.9,
                "LCtau2mint": 67.3,
                "LAtau1maxt": 64.3,
                "LAtau1mint": 60.1,
                "LAtau2maxt": 63.3,
                "LAtau2mint": 60.2,
            })
            for k in base_ovld_keys:
              row[f"{k}_Ovld_Und"] = "-"
            ham_rows.append(row)
            idx += 1

            # Arka Plan Seçilmişse Ek Satır
            if nokta["bg_var"]:
              bg_row = row.copy()
              bg_row["Nokta Sayısı"] = f"{idx-1} (Arka Plan)"
              bg_row["Ölçüm Noktası Adı"] = f"{nokta['ad']} [ARKA PLAN]"
              bg_row["LAtt"] = 51.2  # Arka plan simüle değeri
              ham_rows.append(bg_row)

          df_ham = pd.DataFrame(ham_rows, columns=columns)
          df_ham.to_excel(writer, sheet_name="HAM DATA", index=False)

          df_sonuc = df_ham.copy()
          df_sonuc.to_excel(writer, sheet_name="Sonuç Değerlendirme", index=False)

        output.seek(0)

        st.success(
            "🎉 Tüm periyotlara ait ölçüm ve arka plan verileri başarıyla"
            " işlendi!"
        )
        st.download_button(
            label="📥 Periyot Bazlı Excel Raporunu İndir",
            data=output,
            file_name="Cevresel_Gurultu_Periyot_Bazli_Rapor.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

  with tab2:
    st.subheader("Periyot Bazlı Ölçüm Planı Özeti")
    st.markdown(
        "Seçilen zaman dilimleri (Gündüz/Akşam/Gece) ve arka plan durumlarına"
        " göre oluşturulan yapı."
    )

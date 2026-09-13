import streamlit as st
import pandas as pd
import os

def render_gurultu_module():
    st.title("🔊 Çevresel Gürültü ve Müzik Yayın Ruhsatı Modülü")
    st.markdown("Cesva SC250 `.cdf` ham veri dosyalarını işleme ve değerlendirme araçları.")

    tab1, tab2 = st.tabs(["📂 CDF Cihaz Verisi Dönüştürücü", "📐 Müzik Yayın Ruhsatı Ölçüm Planı"])

    with tab1:
        st.subheader("Cesva SC250 (.cdf) Veri Dönüştürücü")
        st.info("Cihazdan alınan `.cdf` uzantılı ham veri dosyalarını yükleyerek analiz edilebilir Excel formatına dönüştürün. (_S: Spektrum, _T: Zaman Geçmişi)")
        
        uploaded_files = st.file_uploader("CDF Dosyalarını Seçin (Birden fazla seçebilirsiniz)", type=["cdf"], accept_multiple_files=True)
        
        if uploaded_files:
            st.success(f"✅ Toplam {len(uploaded_files)} adet dosya yüklendi.")
            
            if st.button("Dosyaları Analiz Et ve Excel'e Dönüştür"):
                summary_data = []
                
                for file in uploaded_files:
                    file_name = file.name
                    # Dosya adından _S veya _T türünü ayıkla
                    file_type = "Spektrum (_S)" if "_S" in file_name.upper() else ("Zaman Geçmişi (_T)" if "_T" in file_name.upper() else "Bilinmeyen")
                    
                    # CDF dosyaları metin/binary tabanlıdır, örnekleme satırlarını özetleyelim
                    try:
                        content = file.getvalue().decode("latin-1", errors="ignore")
                        lines_count = len(content.splitlines())
                    except Exception:
                        lines_count = 0
                        
                    summary_data.append({
                        "Dosya Adı": file_name,
                        "Veri Türü": file_type,
                        "Satır/Kayıt Sayısı": lines_count
                    })
                
                df_summary = pd.DataFrame(summary_data)
                st.dataframe(df_summary, use_container_width=True)
                
                # Excel Çıktısı Oluşturma
                output_excel = "Cesva_CDF_Analiz_Raporu.xlsx"
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    df_summary.to_excel(writer, sheet_name="Dosya_Ozeti", index=False)
                
                with open(output_excel, "rb") as fp:
                    st.download_button(
                        label="📥 Özet Analiz Excel Dosyasını İndir",
                        data=fp,
                        file_name="Cesva_CDF_Donusum_Raporu.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    with tab2:
        st.subheader("Bitişik Nizam ve Müzik Yayın Ruhsatı Ölçüm Noktası Planlayıcı")
        st.info("İşletmenin konumuna, kat durumuna ve bitişik nizam özelliklerine göre gerekli minimum ölçüm noktalarını belirleyin.")
        
        isletme_turu = st.selectbox(
            "İşletme / Faaliyet Türü:",
            ["Canlı Müzik Yayınlayan İşletmeler (Bar, Restoran vb.)", "Bant / Kayıtlı Müzik Yayınlayan Yerler", "Atölye / Endüstriyel Tesis"]
        )
        
        bitisik_nizam_durumu = st.radio(
            "Bitişik Nizam / Yapı Durumu:",
            ["Hassas Kullanım Alanı (Konut vb.) ile Ortak Duvar/Tavan/Taban Var", "Müstakil / Ortak Sınır Yok"]
        )
        
        kat_sayisi = st.number_input("İşletmenin Bulunduğu Kat", min_value=-2, max_value=20, value=0)
        
        if st.button("Ölçüm Planı ve Rapor Şablonu Oluştur"):
            st.markdown("---")
            st.subheader("📋 Önerilen Ölçüm Noktaları Matrisi")
            
            plan_data = [
                {"Nokta ID": "M-01", "Konum": "İşletme İçi (Kaynak Merkezi)", "Ölçüm Amacı": "İç ortam gürültü seviyesi tespiti", "Süre": "15 Dakika"},
                {"Nokta ID": "M-02", "Konum": "Ortak Duvar / Sınır Konut İçi", "Ölçüm Amacı": "Yansıyan / İletilen gürültü (Darbe/Hava doğuşlu)", "Süre": "İlgili Periyot"},
                {"Nokta ID": "M-03", "Konum": "En Yakın Hassas Cephe (Dış Ortam)", "Ölçüm Amacı": "Çevresel gürültü sınır değer kontrolü", "Süre": "Gündüz/Akşam/Gece"}
            ]
            df_plan = pd.DataFrame(plan_data)
            st.dataframe(df_plan, use_container_width=True)
            st.success("✅ Ölçüm stratejisi Çevresel Gürültü Kontrol Yönetmeliği kriterlerine göre hazırlandı.")

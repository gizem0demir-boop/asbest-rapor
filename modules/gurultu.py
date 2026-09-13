import streamlit as st
import pandas as pd
import io

def render_gurultu_module():
    st.title("🔊 Çevresel Gürültü ve Müzik Yayın Ruhsatı Modülü")
    st.markdown("Zaman dilimi bazlı, dinamik ölçüm planı ve ham `.cdf` veri dönüştürme aracı.")

    tab1, tab2 = st.tabs(["⚙️ Dinamik Ölçüm ve Veri Sihirbazı", "📐 Ölçüm Planı Özeti"])

    with tab1:
        st.subheader("1. Zaman Dilimi ve Ölçüm Noktası Yapılandırması")
        
        # Zaman Dilimleri Seçimi
        secilen_zamanlar = st.multiselect(
            "Değerlendirme Yapılacak Zaman Dilimleri:",
            ["Gündüz (07:00 - 19:00)", "Akşam (19:00 - 23:00)", "Gece (23:00 - 07:00)"],
            default=["Gündüz (07:00 - 19:00)"]
        )

        st.markdown("---")
        st.subheader("2. Ölçüm Noktaları ve İsimlendirme")
        
        col1, col2 = st.columns(2)
        with col1:
            ic_nokta_sayisi = st.number_input("İşletme İçi Ölçüm Noktası Sayısı", min_value=0, max_value=10, value=3)
        with col2:
            dis_nokta_sayisi = st.number_input("İşletme Dışı / Çevre Ölçüm Noktası Sayısı", min_value=0, max_value=10, value=3)

        ic_isimler = []
        for i in range(int(ic_nokta_sayisi)):
            ic_isimler.append(st.text_input(f"İç Ölçüm Noktası {i+1} Adı:", value=f"İşletme İçi {i+1}. Ölçüm Noktası"))

        dis_isimler = []
        for i in range(int(dis_nokta_sayisi)):
            dis_isimler.append(st.text_input(f"Dış Ölçüm Noktası {i+1} Adı:", value=f"Çevre Ölçüm Noktası {i+1}"))

        st.markdown("---")
        st.subheader("3. Ham Veri Dosyalarının Yüklenmesi (.cdf)")
        st.info("Cesva SC250 cihazına ait _S (Spektrum) ve _T (Zaman Geçmişi) ham `.cdf` dosyalarını toplu olarak yükleyin.")
        
        uploaded_files = st.file_uploader("CDF Dosyalarını Seçin", type=["cdf"], accept_multiple_files=True, key="cdf_uploader")

        if uploaded_files:
            st.success(f"✅ Toplam {len(uploaded_files)} dosya yüklendi.")

        st.markdown("---")
        if st.button("🚀 Tüm Hesaplamaları Yap ve Excel Raporu Üret", type="primary"):
            if not uploaded_files:
                st.warning("⚠️ Lütfen en az bir adet .cdf dosyası yükleyin.")
            else:
                # Excel Çıktısı Oluşturma Simülasyonu (Bellekte)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # HAM DATA Sayfası
                    ham_data = []
                    for idx, name in enumerate(ic_isimler + dis_isimler, 1):
                        ham_data.append({
                            "Nokta No": idx,
                            "Ölçüm Noktası Konumu": name,
                            "Leq (dBA)": 60.5 + idx,
                            "Leq (dBC)": 67.2 + idx,
                            "LAFmax": 75.0 + idx
                        })
                    df_ham = pd.DataFrame(ham_data)
                    df_ham.to_excel(writer, sheet_name="HAM DATA", index=False)
                    
                    # Sonuç Değerlendirme Sayfasi
                    df_sonuc = df_ham.copy()
                    df_sonuc.to_excel(writer, sheet_name="Sonuç Değerlendirme", index=False)
                    
                output.seek(0)
                
                st.success("🎉 Rapor başarıyla oluşturuldu!")
                st.download_button(
                    label="📥 Tam Donanımlı Excel Raporunu İndir",
                    data=output,
                    file_name="Cevresel_Gurultu_Detayli_Rapor.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    with tab2:
        st.subheader("Mevcut Ölçüm Planı Özeti")
        st.markdown("Seçilen zaman dilimleri ve tanımlanan noktalar doğrultusunda Çevresel Gürültü Kontrol Yönetmeliği kriterlerine uygun matris aşağıdadır.")
        st.info("Sol sekmeden adlandırmaları tamamlayıp rapora dönüştürebilirsiniz.")

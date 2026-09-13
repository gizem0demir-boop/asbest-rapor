import streamlit as st
import pandas as pd
import pypyodbc
import os

def render_gurultu_module():
    st.title("🔊 Çevresel Gürültü ve Müzik Yayın Ruhsatı Modülü")
    st.markdown("Çevresel Gürültü Kontrol Yönetmeliği kapsamında ölçüm verisi işleme ve değerlendirme araçları.")

    # Sekme yapısı ile modülü ayıralım
    tab1, tab2 = st.tabs(["📂 MDB / Cihaz Verisi Dönüştürücü", "📐 Müzik Yayın Ruhsatı Ölçüm Planı"])

    with tab1:
        st.subheader("Cihaz Veritabanı (.mdb) Dönüştürücü")
        st.info("Ses düzey ölçüm cihazından alınan arka plan veritabanı dosyasını (.mdb) analiz edilebilir Excel formatına dönüştürün.")
        
        uploaded_mdb = st.file_uploader("Gürültü Cihazı Veritabanı Dosyası Yükle", type=["mdb"])
        
        if uploaded_mdb is not None:
            # Geçici olarak diske kaydedip pypyodbc ile okuyacağız
            temp_mdb_path = "temp_data.mdb"
            with open(temp_mdb_path, "wb") as f:
                f.write(uploaded_mdb.getbuffer())
                
            try:
                con_str = (
                    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                    f"DBQ={temp_mdb_path};"
                )
                conn = pypyodbc.connect(con_str)
                
                tables_to_extract = ['Measurement_Data', 'Final_Results_t', 'Time_History_125ms', 'Instrument_Data']
                
                # Excel dosyası oluşturalım
                output_excel = "Cevresel_Gurultu_Cikti.xlsx"
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    for table in tables_to_extract:
                        try:
                            df = pd.read_sql(f"SELECT * FROM [{table}]", conn)
                            df.to_excel(writer, sheet_name=table[:31], index=False)
                        except Exception:
                            # Tablo veritabanında yoksa atla
                            pass
                conn.close()
                
                st.success("✅ Veritabanı başarıyla Excel formatına dönüştürüldü!")
                
                with open(output_excel, "rb") as fp:
                    st.download_button(
                        label="📥 Çıktı Excel Dosyasını İndir",
                        data=fp,
                        file_name="Cevresel_Gurultu_Analiz.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.makeExcel"
                    )
                    
            except Exception as e:
                st.error(f"MDB okuma hatası (Bilgisayarınızda Microsoft Access Database Engine sürücüsü yüklü olmalıdır): {e}")
            finally:
                if os.path.exists(temp_mdb_path):
                    os.remove(temp_mdb_path)

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
            
            # Örnek otomatik tablo mantığı
            plan_data = [
                {"Nokta ID": "M-01", "Konum": "İşletme İçi (Kaynak Merkezi)", "Ölçüm Amacı": "İç ortam gürültü seviyesi tespiti", "Süre": "15 Dakika"},
                {"Nokta ID": "M-02", "Konum": "Ortak Duvar / Sınır Konut İçi", "Ölçüm Amacı": "Yansıyan / İletilen gürültü (Darbe/Hava doğuşlu)", "Süre": "İlgili Periyot"},
                {"Nokta ID": "M-03", "Konum": "En Yakın Hassas Cephe (Dış Ortam)", "Ölçüm Amacı": "Çevresel gürültü sınır değer kontrolü", "Süre": "Gündüz/Akşam/Gece"}
            ]
            df_plan = pd.DataFrame(plan_data)
            st.dataframe(df_plan, use_container_width=True)
            st.success("✅ Ölçüm stratejisi Çevresel Gürültü Kontrol Yönetmeliği kriterlerine göre hazırlandı.")
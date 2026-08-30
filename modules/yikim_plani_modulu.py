import streamlit as st
import pandas as pd

# Sayfa Genişliği
st.set_page_config(page_title="Belge ve Rapor Üretim Paneli", layout="wide")

st.title("📄 Belge ve Rapor Üretim Paneli")

# -------------------------------------------------------------------------
# 1. ADIM: EN ÜSTTE EVRAK TÜRÜ SEÇİMİ
# -------------------------------------------------------------------------
st.markdown("### 📌 Oluşturulacak Evrak / Rapor Türünü Seçin:")
evrak_turu = st.selectbox(
    "-- Seçiniz --",
    [
        "-- Seçiniz --",
        "Müellif - Müteahhit Yıkım Sözleşmesi",
        "Fenni Mesul Taahhütnamesi",
        "Müellif Taahhütnamesi (İdareye Verilecek - Form 2)",
        "Yıkım Planı Raporu (Tam Kapsamlı)"
    ],
    label_visibility="collapsed"
)

# -------------------------------------------------------------------------
# 2. ADIM: TUTANAK / VERİ DOSYASI YÜKLEME
# -------------------------------------------------------------------------
st.markdown("---")
st.markdown("### 📂 Tutanak / Veri Dosyası")
uploaded_file = st.file_uploader("Lütfen tutanak/veri dosyanızı (Excel/CSV) yükleyin:", type=["xlsx", "xls", "csv"])

veriler = {}
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        st.success("Tutanak başarıyla okundu ve yüklendi!")
        # Buraya kendi Excel okuma mantığını (sözlük eşleştirmelerini) ekleyebilirsin.
    except Exception as e:
        st.error(f"Dosya okunurken bir hata oluştu: {e}")

# -------------------------------------------------------------------------
# 3. ADIM: KOŞULLU ALANLAR (Sadece Yıkım Planı Raporu Seçilirse Gözükür)
# -------------------------------------------------------------------------
if evrak_turu == "Yıkım Planı Raporu (Tam Kapsamlı)":
    
    st.markdown("---")
    st.markdown("## 🏗️ Yapı Teknik Özellikleri")
    col1, col2, col3 = st.columns(3)
    with col1:
        toplam_b_bolum = st.text_input("Toplam B. Bölüm Sayısı:")
    with col2:
        toplam_kat = st.text_input("Toplam Kat Sayısı:")
    with col3:
        toplam_alan = st.text_input("Toplam İnşaat Alanı (m²):")

    col4, col5, col6 = st.columns(3)
    with col4:
        nitelik = st.text_input("Niteliği:")
    with col5:
        yapi_sinifi = st.text_input("Yapı Sınıfı / Grubu:")
    with col6:
        bina_yuksekligi = st.text_input("Bina Yüksekliği (m):")

    st.markdown("---")
    st.markdown("## 📐 Yıkım Tekniği, Ekipman ve İş Planı Ayarları")
    yikim_teknigi = st.radio(
        "Yıkım Tekniğini Seçiniz:",
        ["Elle", "Y. Erişimli // Kompakt Makinalı", "Kule ve Diğer Yüksek Erişimli Vinç", 
         "Patlayıcılarla", "Kimyasal Madde Kullanarak", "Sıcak / Metal Tozuyla Kesim", "Diğer"],
        horizontal=True
    )

    st.markdown("### 👥 Yıkımda Görevli Kişi ve Ekipmanlar")
    col7, col8, col9 = st.columns(3)
    with col7:
        personel_sayisi = st.number_input("Personel Sayısı:", value=6)
    with col8:
        makine_sayisi = st.number_input("Makine-Aparat Sayısı:", value=1)
    with col9:
        operator_sayisi = st.number_input("Operatör Sayısı:", value=0)

    p_nitelik_1 = st.checkbox("Personel Nitelikli: İşçi", value=True)
    p_nitelik_2 = st.checkbox("Personel Nitelikli: Uzman", value=True)
    
    makine_turleri = st.text_area("Makine-Aparat Türü:", value="İSKELE\nHAVALI KESME ALETİ\nBALYOZ\nKOREGA BORU")

    col10, col11 = st.columns(2)
    with col10:
        operator_belgesi = st.text_input("Operatör Belgesi / Durumu:", value="GEREKMEZ")
    with col11:
        belge_aciklamasi = st.text_input("Belge Açıklaması / Detayı:", value="ELLE YIKIM ŞARTLARINA UYGUNDUR")

    st.markdown("---")
    st.markdown("## 📋 Yıkım İş Planı ve Nizam Durumu")
    bina_nizam = st.selectbox("Binanın Nizam Durumu:", ["Ayrık Nizam", "Blok Nizam", "Bitişik Nizam"])
    toz_bastirma = st.checkbox("Pulverize Sistemli Toz Bastırma Cihazı Kullanılsın mı?", value=True)

    st.markdown("---")
    st.markdown("## 🗑️ İnşaat ve Yıkıntı Atıkları Miktarları (Ton)")
    atikk_1, atikk_2, atikk_3 = st.columns(3)
    with atikk_1:
        tugla_miktari = st.number_input("Tuğla (17 01 02) Miktarı (Ton):", value=38.00)
    with atikk_2:
        metal_miktari = st.number_input("Karışık Metal (17 04 07) Miktarı (Ton):", value=77.00)
    with atikk_3:
        beton_miktari = st.number_input("Beton (17 01 01) Miktarı (Ton):", value=990.00)

    st.markdown("---")
    st.markdown("## 📷 Yapı Görselleri (Harita Konumu ve Bina Fotoğrafı)")
    gorsel_col1, gorsel_col2 = st.columns(2)
    with gorsel_col1:
        st.file_uploader("Yapının Konumu (Harita Görseli)", type=["png", "jpg", "jpeg"], key="harita")
    with gorsel_col2:
        st.file_uploader("Yapının Fotoğrafı", type=["png", "jpg", "jpeg"], key="bina_foto")

elif evrak_turu != "-- Seçiniz --":
    st.info(f"ℹ️ Seçilen **'{evrak_turu}'** için ayrıntılı yıkım planı girdilerine gerek bulunmamaktadır. Doğrudan rapor bilgileri ile belgeyi üretebilirsiniz.")

# -------------------------------------------------------------------------
# 4. ADIM: ORTAK RAPOR BİLGİLERİ VE BELGE ÜRETİMİ
# -------------------------------------------------------------------------
if evrak_turu != "-- Seçiniz --":
    st.markdown("---")
    st.markdown("## 📄 Rapor Bilgileri ve Belge Üretimi")
    
    rc_col1, rc_col2 = st.columns(2)
    with rc_col1:
        rapor_sayisi = st.text_input("Rapor Sayısı:", value="2026-1276")
    with rc_col2:
        rapor_tarihi = st.date_input("Rapor Tarihi:")

    st.markdown("")
    if st.button("Belgeyi Üret", type="primary"):
        st.success(f"'{evrak_turu}' başarıyla oluşturuldu!")
        # Belge oluşturma / Word (docxtpl) tetikleme kodların buraya gelebilir.
# modules/yikim_plani_modulu.py dosyasının içine eklenecek yapı:

import streamlit as st

def render():
    st.markdown("### Yıkım Planı Raporu Ekranı")
    # Buraya yıkım planı ile ilgili form elemanlarını ve kodlarını koyabilirsin.

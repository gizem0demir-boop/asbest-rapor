import streamlit as st
from modules.asbest import render_asbest_module
from modules.toz import render_toz_module
from modules.ayp import render_ayp_module
from modules.yikim_plani_modulu import render as render_yikim_plani

st.set_page_config(
    page_title="Asya Asbest & Atık Yönetim Sistemi",
    page_icon="🔬",
    layout="wide"
)

# --- Yan Menü (Sidebar) ---
with st.sidebar:
    st.image("https://img.icons8.com/color/96/experimental-copy.png", width=80)
    st.markdown("### 🔬 Laboratuvar Modülü")
    st.write("ASYA Asbest Danışmanlık ve Laboratuvar Hizmetleri Otomasyon Paneli")
    st.markdown("---")
    rapor_turu = st.selectbox(
        "📋 İşlem / Rapor Türü Seçin:",
        [
            "-- Seçiniz --",
            "🔬 Asbest Tür Tayini Raporu",
            "💨 Toz Raporu",
            "♻️ AYP (Atık Yönetim Planı) Raporu",
            "🏗️ Yıkım Planı ve Yasal Evrak Modülü"
        ]
    )
    st.markdown("---")

# --- Ana Ekran Yönlendirmeleri ---
st.title("🏢 Asbest ve Atık Yönetim Rapor Sistemi")
st.markdown("---")

if rapor_turu == "-- Seçiniz --":
    st.warning("⚠️ Lütfen sol menüden oluşturmak istediğiniz **Rapor Türünü** seçin.")
elif rapor_turu == "🔬 Asbest Tür Tayini Raporu":
    render_asbest_module()
elif rapor_turu == "💨 Toz Raporu":
    render_toz_module()
elif rapor_turu == "♻️ AYP (Atık Yönetim Planı) Raporu":
    render_ayp_module()
elif rapor_turu == "🏗️ Yıkım Planı ve Yasal Evrak Modülü":
    render_yikim_plani()

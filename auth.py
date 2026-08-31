import streamlit as st
import pandas as pd
import os

def check_login():
    """kullanicilar.xlsx dosyasından kullanıcı doğrulaması ve oturum yönetimi yapar."""
    
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
        st.session_state["username"] = ""
        st.session_state["role"] = ""

    if not st.session_state["logged_in"]:
        st.markdown("## 🔐 ASYA Otomasyon Sistemi - Güvenli Giriş")
        
        excel_path = "kullanicilar.xlsx"
        if not os.path.exists(excel_path):
            st.error(f"⚠️ '{excel_path}' dosyası bulunamadı! Lütfen ana dizine ekleyin.")
            return False

        try:
            df_users = pd.read_excel(excel_path)
            df_users.columns = [str(c).strip().lower() for c in df_users.columns]
        except Exception as e:
            st.error(f"Excel dosyası okunurken hata oluştu: {e}")
            return False

        with st.form("login_form"):
            username_input = st.text_input("Kullanıcı Adı")
            password_input = st.text_input("Şifre", type="password")
            submit_button = st.form_submit_button("Giriş Yap")
            
            if submit_button:
                user_row = df_users[df_users["kullanici_adi"] == username_input]
                
                if not user_row.empty:
                    gercek_sifre = str(user_row.iloc[0]["sifre"])
                    kullanici_rolu = str(user_row.iloc[0].get("rol", "misafir"))
                    
                    if gercek_sifre == str(password_input):
                        st.session_state["logged_in"] = True
                        st.session_state["username"] = username_input
                        st.session_state["role"] = kullanici_rolu
                        st.success(f"Hoş geldiniz, {username_input}!")
                        st.rerun()
                    else:
                        st.error("Hatalı şifre!")
                else:
                    st.error("Böyle bir kullanıcı adı bulunamadı!")
        return False
    
    return True

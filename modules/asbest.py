import os
from datetime import datetime
import streamlit as st
from docxtpl import DocxTemplate
from docx import Document
from utils import parse_asbest_tutanak, generate_bolum_summary, process_and_get_image

def render_asbest_module():
    st.subheader("🔬 Asbest Katı Numune Analiz Raporu Oluşturucu")

    uploaded_file = st.file_uploader("Numune Tutanağı Excel Dosyasını Yükleyin", type=["xlsx", "xls"], key="asbest_tutanak")

    if uploaded_file is not None:
        info, samples = parse_asbest_tutanak(uploaded_file)
        
        # info veya samples hatalı/boş dönme ihtimaline karşı güvenlik önlemi
        if not isinstance(info, dict):
            info = {}
        if not isinstance(samples, list):
            samples = []

        st.success(f"Tutanak başarıyla okundu! Toplam **{len(samples)}** adet numune tespit edildi.")

        st.markdown("---")
        st.markdown("### 📋 Genel Bilgiler ve Tarih Ayarları")
        
        bugun_tarih = datetime.now().strftime("%d.%m.%Y")

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            musteri_adi = st.text_input("Müşteri / Mal Sahibi:", value=str(info.get("musteri_adi", "")))
            adres = st.text_input("Adres:", value=str(info.get("adres", "")))
            teklif_no = st.text_input("Teklif Numarası:", value=str(info.get("teklif_no", "")))
            numune_tarihi = st.text_input("Numune Alma Tarihi (Tutanaktan):", value=str(info.get("numune_tarihi", "")))
        with col_m2:
            pafta = info.get("pafta", "")
            ada = info.get("ada", "")
            parsel = info.get("parsel", "")
            pafta_ada_parsel = f"{pafta} / {ada} / {parsel}"
            st.info(f"**Pafta / Ada / Parsel:** {pafta_ada_parsel}")
            rapor_tarihi = st.text_input("Rapor Oluşturulma / Yayın Tarihi:", value=bugun_tarih)

        st.markdown("---")
        st.markdown("### 👷 Personel Seçimi")

        numune_nezaret_listesi = [
            "Abdul Samed DEĞİRMENCİ", "Emir UÇARLI", "Ali Kemal DEĞİRMENCİ", 
            "Burak BAYRAKTAR", "Doğucan TAŞTAN", "Emre Can İNEGAZİLİ", 
            "Gözde CANİK", "Furkan TEMİZ", "İsmail AYDIN", "Ogün KAN", "Muharrem YAŞAR"
        ]
        deney_sorumlusu_listesi = [
            "Gizem DEMİR", "Edanur KESGİN", "Ali Kemal DEĞİRMENCİ"
        ]

        col1, col2, col3 = st.columns(3)
        with col1:
            numune_alan = st.selectbox("Numune Alan Personel:", options=numune_nezaret_listesi)
        with col2:
            nezaret_eden = st.selectbox("Nezaret Eden Personel:", options=numune_nezaret_listesi)
        with col3:
            deney_sorumlusu = st.selectbox("Deney Sorumlusu (İmza Yetkilisi):", options=deney_sorumlusu_listesi)

        st.markdown("---")
        st.markdown("### 🖼️ Fotoğraf Yükleme Seçeneği")
        foto_secenegi = st.radio(
            "Rapor fotoğraflarını şimdi yüklemek ister misiniz?",
            ["Fotoğrafları Yükleme (Sonradan Word üzerinde eklenecek)", "Fotoğrafları Şimdi Yükle"],
            horizontal=True
        )

        bina_foto = None
        numune_fotolari = {}

        if foto_secenegi == "Fotoğrafları Şimdi Yükle":
            st.markdown("##### 🏢 Bina / Konut Fotoğrafı")
            bina_foto = st.file_uploader("Bina Dış Görünüş Fotoğrafı", type=["jpg", "jpeg", "png"], key="bina_foto_uploader")

        st.markdown("---")
        st.markdown("### 🔬 Numune Sonuçları ve Bilgileri")

        numuneler = []
        for index, s in enumerate(samples):
            n_kodu = s.get("kod", f"NUM-{index+1}")
            m_turu = s.get("tur", "")
            
            st.markdown(f"**Numune {index+1} | Kod:** `{n_kodu}` | **Malzeme:** `{m_turu}`")
            
            if foto_secenegi == "Fotoğrafları Şimdi Yükle":
                c1, c2, c3 = st.columns([1, 1.5, 1.5])
                with c1:
                    asbest_durumu = st.radio(f"Asbest Durumu ({n_kodu})", ["Yok", "Var"], horizontal=True, key=f"asbest_durum_{index}")
                with c2:
                    if asbest_durumu == "Var":
                        asbest_turu = st.text_input("Tespit Edilen Asbest Türü:", key=f"asbest_tur_{index}")
                        sonuc_metni = f"Asbest tespit edilmiştir ({asbest_turu})" if asbest_turu else "Asbest tespit edilmiştir"
                    else:
                        sonuc_metni = "Asbest tespit edilmedi"
                with c3:
                    numune_fotolari[n_kodu] = st.file_uploader(f"Numune Fotoğrafı ({n_kodu})", type=["jpg", "jpeg", "png"], key=f"foto_upl_{index}")
            else:
                c1, c2 = st.columns([1, 2])
                with c1:
                    asbest_durumu = st.radio(f"Asbest Durumu ({n_kodu})", ["Yok", "Var"], horizontal=True, key=f"asbest_durum_{index}")
                with c2:
                    if asbest_durumu == "Var":
                        asbest_turu = st.text_input("Tespit Edilen Asbest Türü:", key=f"asbest_tur_{index}")
                        sonuc_metni = f"Asbest tespit edilmiştir ({asbest_turu})" if asbest_turu else "Asbest tespit edilmiştir"
                    else:
                        sonuc_metni = "Asbest tespit edilmedi"

            on_islem = "Asitle Muamele" if "marley" in str(m_turu).lower() else "Parçalama"

            numuneler.append({
                "sira": index + 1,
                "tarih": numune_tarihi,
                "kod": n_kodu,
                "tur": m_turu,
                "yer": s.get("yer", ""),
                "yontem": s.get("yontem", ""),
                "strateji": s.get("strateji", ""),
                "homojenite": "Homojen",
                "onislem": on_islem,
                "sonuc": sonuc_metni
            })

        st.markdown("---")
        if st.button("📄 Word Raporunu Oluştur ve İndir", type="primary"):
            try:
                tpl = DocxTemplate("sablon.docx")
                
                context = {
                    "musteri_adi": musteri_adi,
                    "adres": adres,
                    "teklif_no": teklif_no,
                    "pafta": info.get("pafta", ""),
                    "ada": info.get("ada", ""),
                    "parsel": info.get("parsel", ""),
                    "numune_tarihi": numune_tarihi,
                    "rapor_tarihi": rapor_tarihi,
                    "numune_alan": numune_alan,
                    "nezaret_eden": nezaret_eden,
                    "deney_sorumlusu": deney_sorumlusu,
                    "bolum_listesi": generate_bolum_summary(samples)
                }

                if foto_secenegi == "Fotoğrafları Şimdi Yükle":
                    context["bina_foto"] = process_and_get_image(tpl, bina_foto, width_cm=8.0, height_cm=6.0)
                    for index, s in enumerate(samples):
                        n_kodu = s.get("kod", f"NUM-{index+1}")
                        uploaded_img = numune_fotolari.get(n_kodu)
                        context[f"foto_{index+1}"] = process_and_get_image(tpl, uploaded_img, width_cm=6.5, height_cm=5.0)
                else:
                    context["bina_foto"] = ""
                    for index in range(len(samples)):
                        context[f"foto_{index+1}"] = ""

                tpl.render(context)
                temp_path = "gecici_rapor.docx"
                tpl.save(temp_path)

                doc = Document(temp_path)
                target_table = None
                for tbl in doc.tables:
                    if len(tbl.columns) == 10:
                        target_table = tbl
                        break
                if target_table is None:
                    target_table = doc.tables[2]

                if len(target_table.rows) > 2:
                    while len(target_table.rows) > 2:
                        r = target_table.rows[1]._tr
                        r.getparent().remove(r)

                footer_row = target_table.rows[-1]

                for n in numuneler:
                    new_tr = target_table.add_row()._tr
                    footer_row._tr.addprevious(new_tr)
                    new_row_cells = target_table.rows[-2].cells

                    veriler = [
                        str(n["sira"]), str(n["tarih"]), str(n["kod"]), str(n["tur"]),
                        str(n["yer"]), str(n["yontem"]), str(n["strateji"]),
                        str(n["homojenite"]), str(n["onislem"]), str(n["sonuc"])
                    ]
                    for i, val in enumerate(veriler):
                        if i < len(new_row_cells):
                            new_row_cells[i].text = val

                output_path = "cikis_asbest_raporu.docx"
                doc.save(output_path)
                st.success("Rapor başarıyla oluşturuldu!")

                with open(output_path, "rb") as file:
                    st.download_button(
                        label="📥 Oluşturulan Raporu İndir (.docx)",
                        data=file,
                        file_name=f"Asbest_Analiz_Raporu_{teklif_no}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"Hata: {e}")

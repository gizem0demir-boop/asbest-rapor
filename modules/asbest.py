import streamlit as st
import pandas as pd
import openpyxl
import io
import re
from datetime import datetime
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def parse_asbest_tutanak(file_source):
    wb = openpyxl.load_workbook(file_source, data_only=True)
    sheet = wb.active

    # 1. Hücre matrisini oluştur
    cell_matrix = []
    for r in range(1, 40):
        row_vals = []
        for c in range(1, 20):
            val = sheet.cell(row=r, column=c).value
            row_vals.append(str(val).strip() if val is not None else "")
        cell_matrix.append(row_vals)

    # Genel Bilgiler Okuma
    talep_no, numune_tarihi, firma_adi, adres = "", "", "", ""
    pafta, ada, parsel = "-", "-", "-"

    for row in cell_matrix:
        row_str = " ".join(row)
        
        # Talep No ve Tarih (Genelde Row 3-4)
        if "26-" in row_str or "25-" in row_str:
            for cell in row:
                if re.match(r'^\d{2}-\d{2}-\d+', cell):
                    talep_no = cell
                elif re.match(r'^\d{2}\.\d{2}\.\d{4}$', cell):
                    numune_tarihi = cell

        # Firma Adı & Adres
        for cell in row:
            if cell.startswith("Firma Adı:"):
                firma_adi = cell.replace("Firma Adı:", "").strip()
            elif cell.startswith("Firma Adresi:"):
                adres = cell.replace("Firma Adresi:", "").strip()
            elif cell.startswith("Pafta No:"):
                val = cell.replace("Pafta No:", "").strip()
                if val: pafta = val
            elif cell.startswith("Ada No:"):
                val = cell.replace("Ada No:", "").strip()
                if val: ada = val
            elif cell.startswith("Parsel No:"):
                val = cell.replace("Parsel No:", "").strip()
                if val: parsel = val

    if not numune_tarihi:
        numune_tarihi = datetime.now().strftime('%d.%m.%Y')

    # Numune Listesi Ayrıştırma (Sadece verisi olan gerçek numuneler)
    samples = []
    # Satır 10'dan başlayarak başlık satırını pas geçiyoruz
    for r in range(9, len(cell_matrix)):
        row = cell_matrix[r]
        code = row[1] if len(row) > 1 else ""
        tur = row[4] if len(row) > 4 else ""
        yer = row[7] if len(row) > 7 else ""
        bolum = row[10] if len(row) > 10 else ""

        # Filtre: Başlık kelimelerini ve içeriği boş olan satırları kesinlikle alma
        if code and code.lower() not in ["numune numarasi", "numune kodu", "kod", "none", ""]:
            # Eğer malzeme türü, alındığı yer ve bölüm hepsi boşsa pas geç (11. numuneden sonrası)
            if not tur and not yer and not bolum:
                continue
            
            tur_str = tur if tur else "-"
            yer_str = yer
            bolum_str = bolum

            if yer_str and bolum_str:
                tam_yer = f"{yer_str} / {bolum_str}" if yer_str != bolum_str else yer_str
            elif yer_str:
                tam_yer = yer_str
            else:
                tam_yer = bolum_str or "-"

            on_islem = "Asitle Muamele" if "marley" in tur_str.lower() else "Parçalama"

            samples.append({
                'kod': code,
                'tur': tur_str,
                'yer': tam_yer,
                'yontem': 'TS EN ISO 16000-7',
                'strateji': 'Görsel ve Alansal',
                'homojenite': 'Homojen',
                'onislem': on_islem
            })

    info = {
        'talep_no': talep_no,
        'numune_tarihi': numune_tarihi,
        'firma_adi': firma_adi,
        'adres': adres,
        'pafta': pafta,
        'ada': ada,
        'parsel': parsel
    }

    return info, samples

def generate_word_report(info, samples, person_alan, person_nezaret, person_deney):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("ASBEST KATI NUMUNE ANALİZ RAPORU")
    run_title.font.bold = True
    run_title.font.size = Pt(16)
    run_title.font.name = "Arial"
    run_title.font.color.rgb = RGBColor(0, 51, 102)

    doc.add_paragraph()

    table_info = doc.add_table(rows=6, cols=2)
    table_info.style = 'Table Grid'

    pafta_ada_parsel = f"{info.get('pafta', '-') } / {info.get('ada', '-')} / {info.get('parsel', '-')}"

    info_data = [
        ("Müşteri / Firma Adı:", info.get('firma_adi', '') or "-"),
        ("Adres:", info.get('adres', '') or "-"),
        ("Teklif / Talep No:", info.get('talep_no', '') or "-"),
        ("Pafta / Ada / Parsel:", pafta_ada_parsel),
        ("Numune Alma Tarihi:", info.get('numune_tarihi', '')),
        ("Rapor Tarihi:", datetime.now().strftime("%d.%m.%Y"))
    ]

    for idx, (label, val) in enumerate(info_data):
        row_cells = table_info.rows[idx].cells
        row_cells[0].text = label
        row_cells[0].paragraphs[0].runs[0].font.bold = True
        row_cells[0].paragraphs[0].runs[0].font.name = "Arial"
        row_cells[0].paragraphs[0].runs[0].font.size = Pt(10)
        
        row_cells[1].text = str(val)
        row_cells[1].paragraphs[0].runs[0].font.name = "Arial"
        row_cells[1].paragraphs[0].runs[0].font.size = Pt(10)

    doc.add_paragraph()

    p_h2 = doc.add_paragraph()
    run_h2 = p_h2.add_run("1. KATI NUMUNE ANALİZ SONUÇLARI")
    run_h2.font.bold = True
    run_h2.font.size = Pt(12)
    run_h2.font.name = "Arial"
    run_h2.font.color.rgb = RGBColor(0, 51, 102)

    headers = [
        "Sıra", "Numune Kodu", "Malzeme Türü", "Alındığı Yer / Bölüm", 
        "Analiz Yöntemi", "Homojenlik", "Ön İşlem", "Analiz Sonucu"
    ]

    table_samples = doc.add_table(rows=len(samples) + 1, cols=len(headers))
    table_samples.style = 'Table Grid'

    hdr_cells = table_samples.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p.runs:
            p.runs[0].font.bold = True
            p.runs[0].font.size = Pt(9)
            p.runs[0].font.name = "Arial"

    for idx, s in enumerate(samples):
        row_cells = table_samples.rows[idx + 1].cells
        vals = [
            str(idx + 1), s['kod'], s['tur'], s['yer'], 
            s['yontem'], s['homojenite'], s['onislem'], s.get('sonuc', 'Asbest tespit edilmedi')
        ]
        for i, val in enumerate(vals):
            row_cells[i].text = val
            p = row_cells[i].paragraphs[0]
            if i in [0, 1, 4, 5, 6]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                p.runs[0].font.size = Pt(9)
                p.runs[0].font.name = "Arial"

    doc.add_paragraph()

    p_app = doc.add_paragraph()
    run_app = p_app.add_run("2. ONAY VE İMZA BİLGİLERİ")
    run_app.font.bold = True
    run_app.font.size = Pt(12)
    run_app.font.name = "Arial"
    run_app.font.color.rgb = RGBColor(0, 51, 102)

    table_sig = doc.add_table(rows=2, cols=3)
    table_sig.style = 'Table Grid'

    sig_titles = ["Numune Alan Personel", "Nezaret Eden Sorumlu", "Deney Sorumlusu (İmza)"]
    for i, title in enumerate(sig_titles):
        cell = table_sig.rows[0].cells[i]
        cell.text = title
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p.runs:
            p.runs[0].font.bold = True
            p.runs[0].font.size = Pt(10)
            p.runs[0].font.name = "Arial"

    sig_names = [person_alan, person_nezaret, person_deney]
    for i, name in enumerate(sig_names):
        cell = table_sig.rows[1].cells[i]
        cell.text = f"\n\n{name}\n"
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p.runs:
            p.runs[0].font.size = Pt(10)
            p.runs[0].font.name = "Arial"

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def render_asbest_module():
    st.title("🧪 Asbest Katı Numunesi Raporlama Otomasyonu")

    uploaded_file = st.file_uploader("Numune Tutanağı Excel Dosyasını Yükleyin (.xlsx)", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            info, samples = parse_asbest_tutanak(uploaded_file)
            st.success(f"Tutanak okundu! Toplam **{len(samples)}** adet geçerli numune bulundu.")

            st.markdown("---")
            st.subheader("📋 Genel Bilgiler")

            col1, col2 = st.columns(2)

            with col1:
                info['firma_adi'] = st.text_input("Müşteri / Firma Adı", value=info.get('firma_adi', ''))
                info['adres'] = st.text_input("Adres", value=info.get('adres', ''))
                info['talep_no'] = st.text_input("Teklif / Talep No", value=info.get('talep_no', ''))

            with col2:
                info['pafta'] = st.text_input("Pafta No", value=info.get('pafta', ''))
                info['ada'] = st.text_input("Ada No", value=info.get('ada', ''))
                info['parsel'] = st.text_input("Parsel No", value=info.get('parsel', ''))

            st.markdown("---")
            st.subheader("👤 Personel Seçimi")
            
            personel_listesi = [
                "Abdul Samed DEĞİRMENCİ", "Emir UÇARLI", "Ali Kemal DEĞİRMENCİ", 
                "Burak BAYRAKTAR", "Doğucan TAŞTAN", "Emre Can İNEGAZİLİ", 
                "Gözde CANİK", "Furkan TEMİZ", "İsmail AYDIN", "Ogün KAN", "Muharrem YAŞAR"
            ]
            deney_listesi = ["Gizem DEMİR", "Edanur KESGİN", "Ali Kemal DEĞİRMENCİ"]

            col_p1, col_p2, col_p3 = st.columns(3)
            with col_p1:
                person_alan = st.selectbox("Numune Alan Personel:", personel_listesi, index=0)
            with col_p2:
                person_nezaret = st.selectbox("Nezaret Eden Personel:", personel_listesi, index=10)
            with col_p3:
                person_deney = st.selectbox("Deney Sorumlusu:", deney_listesi, index=0)

            st.markdown("---")
            st.subheader("🖼️ Fotoğraf Yükleme Seçeneği")

            foto_option = st.radio(
                "Rapor fotoğraflarını şimdi yüklemek ister misiniz?",
                ["Fotoğrafları Yükleme (Sonradan Word üzerinde eklenecek)", "Fotoğrafları Şimdi Yükle"],
                horizontal=True
            )

            bina_fotolari = {}
            numune_fotolari = {}

            if foto_option == "Fotoğrafları Şimdi Yükle":
                st.markdown("---")
                st.subheader("🏢 Bina / Konut Fotoğrafları")
                
                b_col1, b_col2, b_col3 = st.columns(3)
                with b_col1:
                    bina_fotolari['bina_1'] = st.file_uploader("Bina Dış Görünüş 1", type=["jpg", "jpeg", "png"], key="bina_1")
                with b_col2:
                    bina_fotolari['bina_2'] = st.file_uploader("Bina Dış Görünüş 2", type=["jpg", "jpeg", "png"], key="bina_2")
                with b_col3:
                    bina_fotolari['bina_3'] = st.file_uploader("Bina Dış Görünüş 3", type=["jpg", "jpeg", "png"], key="bina_3")

            st.markdown("---")
            st.subheader("🧪 Numune Analiz Sonuçları ve Fotoğrafları")

            for idx, s in enumerate(samples):
                st.markdown(f"**Numune {idx+1}:** `{s['kod']}` - **Malzeme:** `{s['tur']}` ({s['yer']})")
                
                if foto_option == "Fotoğrafları Şimdi Yükle":
                    col_analiz, col_fotos = st.columns([1, 2])
                    
                    with col_analiz:
                        asbest_var_mi = st.radio(f"Asbest Durumu ({s['kod']})", ["Yok", "Var"], horizontal=True, key=f"rad_{idx}")
                        if asbest_var_mi == "Var":
                            tur_input = st.text_input(f"Asbest Türü ({s['kod']})", placeholder="Örn: Krizotil", key=f"txt_{idx}")
                            s['sonuc'] = f"Asbest tespit edilmiştir ({tur_input})" if tur_input else "Asbest tespit edilmiştir"
                        else:
                            s['sonuc'] = "Asbest tespit edilmedi"

                    with col_fotos:
                        st.caption(f"📸 {s['kod']} Fotoğrafları")
                        img_col1, img_col2, img_col3 = st.columns(3)
                        with img_col1:
                            f_uzak = st.file_uploader("Uzak Çekim", type=["jpg", "jpeg", "png"], key=f"uzak_{idx}")
                        with img_col2:
                            f_yakin = st.file_uploader("Yakın Çekim", type=["jpg", "jpeg", "png"], key=f"yakin_{idx}")
                        with img_col3:
                            f_poset = st.file_uploader("Poşetli Hali", type=["jpg", "jpeg", "png"], key=f"poset_{idx}")
                        
                        numune_fotolari[s['kod']] = {
                            'uzak': f_uzak,
                            'yakin': f_yakin,
                            'poset': f_poset
                        }
                else:
                    res_col1, res_col2 = st.columns([1, 2])
                    with res_col1:
                        asbest_var_mi = st.radio(f"Asbest Durumu ({s['kod']})", ["Yok", "Var"], horizontal=True, key=f"rad_{idx}")
                    with res_col2:
                        if asbest_var_mi == "Var":
                            tur_input = st.text_input(f"Asbest Türü ({s['kod']})", placeholder="Örn: Krizotil", key=f"txt_{idx}")
                            s['sonuc'] = f"Asbest tespit edilmiştir ({tur_input})" if tur_input else "Asbest tespit edilmiştir"
                        else:
                            s['sonuc'] = "Asbest tespit edilmedi"
                
                st.markdown("---")

            if st.button("📄 Word Raporunu Oluştur", type="primary"):
                word_bytes = generate_word_report(info, samples, person_alan, person_nezaret, person_deney)
                st.success("Rapor oluşturuldu!")
                st.download_button(
                    label="💾 Word Raporunu İndir (.docx)",
                    data=word_bytes,
                    file_name=f"Asbest_Analiz_Raporu_{info['talep_no'] or 'Rapor'}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error(f"Hata oluştu: {e}")

import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def parse_asbest_tutanak(file_source):
    wb = openpyxl.load_workbook(file_source, data_only=True)
    sheet = wb.active

    # Yardımcı arama fonksiyonu: Satırdaki tüm hücreleri tarar ve etiketi bulunca değerini döndürür
    def find_field_value(start_row, end_row, keyword):
        for r in range(start_row, end_row + 1):
            for c in range(1, 15):
                val = sheet.cell(row=r, column=c).value
                if val and keyword.lower() in str(val).lower():
                    # Hücre tek başına "Firma Adı: ABC" şeklinde de olabilir, "Firma Adı:" yan hücrede de olabilir
                    text = str(val)
                    if ":" in text and len(text.split(":", 1)[1].strip()) > 0:
                        return text.split(":", 1)[1].strip()
                    # Yanındaki dolu hücreye bak
                    for next_c in range(c + 1, c + 5):
                        next_val = sheet.cell(row=r, column=next_c).value
                        if next_val is not None and str(next_val).strip() != "":
                            return str(next_val).strip()
        return ""

    # Üst Bilgileri Esnek Çekme
    talep_no = find_field_value(1, 5, "Teklif") or find_field_value(1, 5, "Talep")
    firma_adi = find_field_value(4, 7, "Firma Adı") or find_field_value(4, 7, "Müşteri")
    adres = find_field_value(5, 8, "Adres")
    pafta = find_field_value(6, 9, "Pafta")
    ada = find_field_value(6, 9, "Ada")
    parsel = find_field_value(6, 9, "Parsel")

    # Tarih Çekme
    numune_tarihi = ""
    for r in range(1, 6):
        for c in range(1, 15):
            val = sheet.cell(row=r, column=c).value
            if val and "tarih" in str(val).lower():
                if hasattr(val, 'strftime'):
                    numune_tarihi = val.strftime('%d.%m.%Y')
                elif ":" in str(val):
                    numune_tarihi = str(val).split(":")[-1].strip()
                break

    if not numune_tarihi:
        numune_tarihi = datetime.now().strftime('%d.%m.%Y')

    samples = []
    # Numune satırları (10-25 arası veriler)
    for r in range(10, 26):
        code = sheet.cell(row=r, column=2).value
        tur = sheet.cell(row=r, column=5).value
        yer = sheet.cell(row=r, column=8).value
        bolum = sheet.cell(row=r, column=11).value

        if code and (tur or yer or bolum):
            tur_str = str(tur).strip() if tur else "-"
            yer_str = str(yer).strip() if yer else ""
            bolum_str = str(bolum).strip() if bolum else ""

            if yer_str and bolum_str:
                tam_yer = f"{yer_str} / {bolum_str}"
            elif yer_str:
                tam_yer = yer_str
            else:
                tam_yer = bolum_str or "-"

            on_islem = "Asitle Muamele" if "marley" in tur_str.lower() else "Parçalama"

            samples.append({
                'kod': str(code).strip(),
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

    pafta_ada_parsel = f"{info['pafta']} / {info['ada']} / {info['parsel']}".strip(" /")

    info_data = [
        ("Müşteri / Firma Adı:", info['firma_adi'] or "-"),
        ("Adres:", info['adres'] or "-"),
        ("Teklif / Talep No:", info['talep_no'] or "-"),
        ("Pafta / Ada / Parsel:", pafta_ada_parsel or "-"),
        ("Numune Alma Tarihi:", info['numune_tarihi']),
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
            s['yontem'], s['homojenite'], s['onislem'], s['sonuc']
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
    """app.py tarafından çağrılan ana modül fonksiyonu"""
    st.title("🧪 Asbest Katı Numunesi Raporlama Otomasyonu")

    uploaded_file = st.file_uploader("Numune Tutanağı Excel Dosyasını Yükleyin (.xlsx)", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            info, samples = parse_asbest_tutanak(uploaded_file)
            st.success(f"Tutanak okundu! Toplam **{len(samples)}** adet geçerli numune bulundu.")

            st.markdown("---")
            st.subheader("📋 Genel Bilgiler")
            
            c1, c2 = st.columns(2)
            with c1:
                info['firma_adi'] = st.text_input("Müşteri / Firma Adı", value=info['firma_adi'])
                info['adres'] = st.text_input("Adres", value=info['adres'])
                info['talep_no'] = st.text_input("Teklif / Talep No", value=info['talep_no'])
            with c2:
                info['pafta'] = st.text_input("Pafta No", value=info['pafta'])
                info['ada'] = st.text_input("Ada No", value=info['ada'])
                info['parsel'] = st.text_input("Parsel No", value=info['parsel'])

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
            st.subheader("🧪 Numune Analiz Sonuçları")

            for idx, s in enumerate(samples):
                st.markdown(f"**Numune {idx+1}:** `{s['kod']}` - **Malzeme:** `{s['tur']}` ({s['yer']})")
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
                    file_name=f"Asbest_Analiz_Raporu_{info['talep_no']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error(f"Hata oluştu: {e}")

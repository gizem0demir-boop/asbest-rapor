import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Tutanağın üst bilgi ve numune tablosunu okuyan fonksiyon
def parse_asbest_tutanak(file):
    df_raw = pd.read_excel(file, header=None)
    
    info = {
        'musteri_adi': 'ABC İnşaat',
        'adres': '-',
        'pafta': '-',
        'ada': '-',
        'parsel': '-',
        'numune_tarihi': '20.08.2026',
        'teklif_no': '26-08-5191',
        'telefon': '-'
    }
    
    for idx in range(min(10, len(df_raw))):
        row_values = [str(x) for x in df_raw.iloc[idx].values if pd.notna(x)]
        row_text = " ".join(row_values)
        
        if "Talep Numarası" in row_text:
            if idx + 1 < len(df_raw):
                val = str(df_raw.iloc[idx+1].values[0])
                if val and val != "nan":
                    info['teklif_no'] = val.strip()

        if "Firma Adı:" in row_text:
            m = re.search(r'Firma Adı:\s*(.*?)(?:Telefon|$)', row_text)
            if m and m.group(1).strip():
                info['musteri_adi'] = m.group(1).strip()
        
        if "Telefon Numarası:" in row_text:
            m = re.search(r'Telefon Numarası:\s*(.*)', row_text)
            if m and m.group(1).strip():
                info['telefon'] = m.group(1).strip()

        if "Firma Adresi:" in row_text:
            m = re.search(r'Firma Adresi:\s*(.*)', row_text)
            if m and m.group(1).strip():
                info['adres'] = m.group(1).strip()
                
        if "Pafta No:" in row_text or "Parsel No:" in row_text:
            p = re.search(r'Pafta\s*No:\s*([^\s|]*)(?=\s*Ada|$)', row_text, re.IGNORECASE)
            a = re.search(r'Ada\s*No:\s*([^\s|]*)(?=\s*Parsel|$)', row_text, re.IGNORECASE)
            pr = re.search(r'Parsel\s*No:\s*([^\s|]*)(?=$)', row_text, re.IGNORECASE)
            
            if p and p.group(1).strip(): info['pafta'] = p.group(1).strip()
            if a and a.group(1).strip(): info['ada'] = a.group(1).strip()
            if pr and pr.group(1).strip(): info['parsel'] = pr.group(1).strip()

        if "Tarih" in row_text:
            for cell in row_values:
                if re.match(r'\d{2}\.\d{2}\.\d{4}', cell):
                    info['numune_tarihi'] = cell

    samples = []
    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
        
        code_match = re.search(r'NK\.\d+\.\d+-\d+', row_str)
        if code_match:
            code = code_match.group(0)
            non_empty = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != '']
            
            if len(non_empty) >= 3 and any(k in non_empty[1] for k in ['NK.', 'NK']):
                tur = non_empty[2] if len(non_empty) > 2 else "Beton / Sıva"
                yer = non_empty[3] if len(non_empty) > 3 else "-"
                yontem = non_empty[4] if len(non_empty) > 4 else "-"
                strateji = non_empty[5] if len(non_empty) > 5 else "-"
                
                samples.append({
                    'kod': code,
                    'tur': tur,
                    'yer': yer,
                    'yontem': yontem,
                    'strateji': strateji
                })

    return info, samples

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

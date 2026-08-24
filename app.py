import os
import pandas as pd
from flask import Flask, render_template, request, send_file
from docx import Document

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def calculate_ayp_excel(file_path):
    """Ayp Hesaplama.xls dosyasından tüm atık ve hesaplama verilerini okur"""
    xls = pd.ExcelFile(file_path)
    df1 = pd.read_excel(xls, sheet_name='Sayfa1', header=None)
    df2 = pd.read_excel(xls, sheet_name='Sayfa2', header=None)
    
    # Sayfa1 ve Sayfa2'deki hesaplanan değerleri güvenli bir şekilde çekiyoruz
    def get_val(df, row, col, default=0):
        try:
            val = df.iloc[row, col]
            return val if pd.notna(val) else default
        except:
            return default

    # Örnek hücre okumaları (Ayp Hesaplama.xls yapılandırmanıza göre)
    hesaplar = {
        "TUĞLA_MIKTARI": get_val(df1, 6, 10, 9504),
        "ALÇI_MIKTARI": get_val(df1, 10, 9, 31680),
        "BETON_MIKTARI": get_val(df1, 15, 9, 177120),
        "ATERMIT_HESAP_DETAYI": get_val(df1, 53, 7, 0), # Veya satır/kolon indeksine göre
        "AHŞAP_MIKTARI": get_val(df1, 25, 7, 345.6),
        "SERAMİK_MIKTARI": get_val(df1, 32, 5, 5174.1),
        "KİREMİT_MIKTARI": get_val(df1, 27, 4, 3690),
        "DEMİR_HESAP_DETAYI": get_val(df1, 51, 5, 13120),
        "KAĞIT_HESAP_DETAYI": 12, # Kağıt toplamı
    }
    
    # Sayfa2'deki özet tablodan da alternatif okuma yapabilirsiniz
    for index, row in df2.iterrows():
        atik_adi = str(row[5]) if len(row) > 5 else ""
        miktar = row[6] if len(row) > 6 else None
        if pd.notna(miktar):
            if "tuğla" in atik_adi.lower():
                hesaplar["TUĞLA_MIKTARI"] = miktar
            elif "beton" in atik_adi.lower():
                hesaplar["BETON_MIKTARI"] = miktar

    return hesaplar

def replace_tags_in_paragraph(paragraph, data_dict):
    """Paragraf içindeki etiketleri bulur ve değerleriyle değiştirir (biçimlendirmeyi bozmadan)"""
    for key, value in data_dict.items():
        tag = f"{{{{{key}}}}}"
        if tag in paragraph.text:
            for run in paragraph.runs:
                if tag in run.text:
                    run.text = run.text.replace(tag, str(value))
            # Eğer run'lara bölünmüşse direkt paragraf text üzerinden de güncelleyelim
            if tag in paragraph.text:
                paragraph.text = paragraph.text.replace(tag, str(value))

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        rapor_turu = request.form.get('rapor_turu')
        
        if rapor_turu == 'ayp':
            tutanak_file = request.files.get('tutanak_file')
            ayp_excel_file = request.files.get('ayp_excel_file')
            
            if tutanak_file and ayp_excel_file:
                tutanak_path = os.path.join(UPLOAD_FOLDER, tutanak_file.filename)
                excel_path = os.path.join(UPLOAD_FOLDER, ayp_excel_file.filename)
                
                tutanak_file.save(tutanak_path)
                ayp_excel_file.save(excel_path)
                
                # Excel'den hesaplamaları sözlük olarak al
                hesaplanan_degerler = calculate_ayp_excel(excel_path)
                
                # sablon_ayp.docx dosyasını yükle
                doc = Document('sablon_ayp.docx')
                
                # Paragraflardaki etiketleri değiştir
                for paragraph in doc.paragraphs:
                    replace_tags_in_paragraph(paragraph, hesaplanan_degerler)
                
                # Tablolardaki etiketleri değiştir
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                replace_tags_in_paragraph(paragraph, hesaplanan_degerler)
                
                output_path = os.path.join(UPLOAD_FOLDER, 'AYP_Raporu_Cikti.docx')
                doc.save(output_path)
                return send_file(output_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

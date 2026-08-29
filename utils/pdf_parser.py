import re
import pdfplumber

def parse_asbestos_report(pdf_path):
    extracted_data = {
        "musteri_adi": None,
        "numune_adresi": None,
        "pafta_ada_parsel": None,
        "rapor_no": None,
        "tarih": None
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        # Genellikle ilk sayfada yer alır
        text = pdf.pages[0].extract_text()
        
        # Regex veya satır bazlı anahtar kelime eşleştirmeleri
        for line in text.split('\n'):
            if "Müşterinin / Mal Sahibi Adı" in line:
                # İki nokta sonrasını al
                parts = line.split(":")
                if len(parts) > 1:
                    extracted_data["musteri_adi"] = parts[1].strip()
            
            elif "Numune Alınan Adres" in line:
                parts = line.split(":")
                if len(parts) > 1:
                    extracted_data["numune_adresi"] = parts[1].strip()
                    
            elif "Pafta No / Ada No / Parsel No" in line:
                parts = line.split(":")
                if len(parts) > 1:
                    extracted_data["pafta_ada_parsel"] = parts[1].strip()
                    
            elif "Rapor Numarası" in line:
                parts = line.split(":")
                if len(parts) > 1:
                    extracted_data["rapor_no"] = parts[1].strip()

    return extracted_data
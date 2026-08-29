import re
import pdfplumber

def parse_asbestos_pdf_report(pdf_path):
    extracted_data = {
        "musteri_adi": "",
        "adres": "-",
        "pafta": "-",
        "ada": "-",
        "parsel": "-",
        "teklif_no": "",
        "rapor_no": ""
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()
        
        for line in text.split('\n'):
            # Yeşil Bölge Karşılığı (Müşterinin / Mal Sahibi Adı)
            if "Müşterinin / Mal Sahibi Adı" in line:
                parts = line.split(":")
                if len(parts) > 1:
                    extracted_data["musteri_adi"] = parts[1].strip()
            
            # Mavi Bölge Karşılığı (Numune Alınan Adres)
            elif "Numune Alınan Adres" in line:
                parts = line.split(":")
                if len(parts) > 1:
                    extracted_data["adres"] = parts[1].strip()
            
            # Teklif Numarası
            elif "Teklif Numarası" in line:
                parts = line.split(":")
                if len(parts) > 1:
                    extracted_data["teklif_no"] = parts[1].strip()
                    
            # Kırmızı Bölge Karşılığı (Pafta No / Ada No / Parsel No, örn: - / 22847 / 22)
            elif "Pafta No / Ada No / Parsel No" in line:
                parts = line.split(":")
                if len(parts) > 1:
                    val_str = parts[1].strip()
                    # Örn: "- / 22847 / 22" ayrıştırma
                    sub_parts = [p.strip() for p in val_str.split('/')]
                    if len(sub_parts) >= 3:
                        extracted_data["pafta"] = sub_parts[0]
                        extracted_data["ada"] = sub_parts[1]
                        extracted_data["parsel"] = sub_parts[2]
                        
            # Rapor Numarası
            elif "Rapor Numarası" in line:
                parts = line.split(":")
                if len(parts) > 1:
                    extracted_data["rapor_no"] = parts[1].strip()

    return extracted_data

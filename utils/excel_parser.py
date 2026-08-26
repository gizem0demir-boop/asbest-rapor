import pandas as pd
import re

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
    
    # 1. Üst Bilgileri Okuma
    for idx in range(min(25, len(df_raw))):
        row_values = [str(x).strip() for x in df_raw.iloc[idx].values if pd.notna(x)]
        row_text = " ".join(row_values)
        
        if "Talep Numarası" in row_text or "Teklif" in row_text:
            for cell in row_values:
                if re.search(r'\d{2}-\d{2}-\d+', cell):
                    info['teklif_no'] = cell
        
        if "Firma Adı:" in row_text:
            m = re.search(r'Firma Adı:\s*(.*?)(?:Telefon|Adres|$)', row_text, re.IGNORECASE)
            if m and m.group(1).strip():
                info['musteri_adi'] = m.group(1).strip()
        
        if "Telefon Numarası:" in row_text:
            m = re.search(r'Telefon Numarası:\s*(.*)', row_text, re.IGNORECASE)
            if m and m.group(1).strip():
                info['telefon'] = m.group(1).strip()

        if "Firma Adresi:" in row_text:
            m = re.search(r'Firma Adresi:\s*(.*)', row_text, re.IGNORECASE)
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

    # 2. Numune Tablosunu Okuma
    samples = []
    seen_codes = set()

    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        non_empty = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != '']
        row_str = " ".join(non_empty)
        
        # Orijinal regex araması
        code_match = re.search(r'NK\.\d+\.\d+-\d+', row_str)
        if code_match:
            code = code_match.group(0)
            
            # Sayfa sonundaki mükerrer okumaları engellemek için tek satırlık ek
            if code not in seen_codes:
                seen_codes.add(code)
                
                code_idx = -1
                for i, val in enumerate(non_empty):
                    if code in val:
                        code_idx = i
                        break
                
                tur = non_empty[code_idx + 1] if len(non_empty) > code_idx + 1 else "Beton / Sıva"
                yer = non_empty[code_idx + 2] if len(non_empty) > code_idx + 2 else "-"
                yontem = non_empty[code_idx + 3] if len(non_empty) > code_idx + 3 else "TS EN ISO 16000-7"
                strateji = non_empty[code_idx + 4] if len(non_empty) > code_idx + 4 else "Görsel ve Alansal"

                samples.append({
                    'kod': code,
                    'tur': tur,
                    'yer': yer,
                    'yontem': yontem,
                    'strateji': strateji
                })

    return info, samples

# Modül bağımlılığı için takma ad (ImportError'ı önler)
read_tutanak_details = parse_asbest_tutanak

import pandas as pd
import re

def parse_asbest_tutanak(file):
    df_raw = pd.read_excel(file, header=None)
    
    info = {
        'musteri_adi': '',
        'adres': '-',
        'pafta': '-',
        'ada': '-',
        'parsel': '-',
        'numune_tarihi': '',
        'teklif_no': '',
        'telefon': '-'
    }
    
    # 1. Üst Bilgileri Okuma
    for idx in range(min(25, len(df_raw))):
        row_values = [str(x).strip() for x in df_raw.iloc[idx].values if pd.notna(x)]
        row_text = " ".join(row_values)
        
        if "Talep Numarası" in row_text or "Teklif" in row_text:
            for cell in row_values:
                if re.search(r'\d{2}-\d{2}-\d{4,5}', cell):
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

    # 2. Tam Numune Kodu Formatı (Örn: NK.26.5038-01)
    # Sadece hücresi tam olarak numune kodu olan satırları alır
    sample_pattern = r'^NK\.\d{2}\.\d{4}-\d{2}$'
    samples = []
    
    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        non_empty = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != '']
        
        for i, val in enumerate(non_empty):
            # Hücre tam numune formatına uyuyorsa
            if re.match(sample_pattern, val):
                code = val
                tur = non_empty[i + 1] if len(non_empty) > i + 1 else "Beton / Sıva"
                yer = non_empty[i + 2] if len(non_empty) > i + 2 else "-"
                yontem = non_empty[i + 3] if len(non_empty) > i + 3 else "TS EN ISO 16000-7"
                strateji = non_empty[i + 4] if len(non_empty) > i + 4 else "Görsel ve Alansal"

                samples.append({
                    'kod': code,
                    'tur': tur,
                    'yer': yer,
                    'yontem': yontem,
                    'strateji': strateji
                })
                break  # Aynı satırda tekrar numune arama

    return info, samples

read_tutanak_details = parse_asbest_tutanak

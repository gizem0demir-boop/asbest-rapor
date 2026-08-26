import pandas as pd
import re

def parse_asbest_tutanak(file):
    # Tüm hücreleri string olarak yükle
    df_raw = pd.read_excel(file, header=None, dtype=str)
    
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
    
    samples = []
    seen_codes = set()

    for idx in range(len(df_raw)):
        # Satırdaki boş olmayan hücreleri temizle
        row_cells = [str(x).strip() for x in df_raw.iloc[idx].values if pd.notna(x) and str(x).strip() not in ['', 'nan', 'None']]
        if not row_cells:
            continue
            
        row_str = " ".join(row_cells)

        # ------------------- 1. ÜST BİLGİ OKUMA -------------------
        # Teklif / Talep Numarası
        if "Talep Numarası" in row_str or "Teklif" in row_str:
            for cell in row_cells:
                m = re.search(r'(\d{2}-\d{2}-\d+)', cell)
                if m:
                    info['teklif_no'] = m.group(1)

        # Müşteri Adı
        if "Firma Adı:" in row_str or "Müşteri Adı:" in row_str:
            m = re.search(r'(?:Firma|Müşteri)\s*Adı:\s*(.*?)(?:Telefon|Adres|Pafta|$)', row_str, re.IGNORECASE)
            if m and m.group(1).strip():
                info['musteri_adi'] = m.group(1).strip()

        # Adres
        if "Firma Adresi:" in row_str or "Adres:" in row_str:
            m = re.search(r'(?:Firma Adresi|Adres):\s*(.*?)(?:Pafta|Ada|Parsel|$)', row_str, re.IGNORECASE)
            if m and m.group(1).strip():
                info['adres'] = m.group(1).strip()

        # Pafta / Ada / Parsel
        if any(k in row_str for k in ["Pafta No:", "Ada No:", "Parsel No:"]):
            p = re.search(r'Pafta\s*No:\s*([^\s|]*)(?=\s*Ada|$)', row_str, re.IGNORECASE)
            a = re.search(r'Ada\s*No:\s*([^\s|]*)(?=\s*Parsel|$)', row_str, re.IGNORECASE)
            pr = re.search(r'Parsel\s*No:\s*([^\s|]*)(?=$)', row_str, re.IGNORECASE)
            if p and p.group(1).strip(): info['pafta'] = p.group(1).strip()
            if a and a.group(1).strip(): info['ada'] = a.group(1).strip()
            if pr and pr.group(1).strip(): info['parsel'] = pr.group(1).strip()

        # Numune Tarihi (GG.AA.YYYY)
        if "Tarih" in row_str:
            for cell in row_cells:
                m = re.search(r'(\d{2}\.\d{2}\.\d{4})', cell)
                if m:
                    info['numune_tarihi'] = m.group(1)

        # ------------------- 2. NUMUNE TABLOSU OKUMA -------------------
        # NK. ile başlayan her türlü numune kodunu yakalar (Esnek Regex)
        code_match = re.search(r'NK\.[\w\.\-]+', row_str)
        if code_match:
            code = code_match.group(0).rstrip('.')
            
            # Sadece geçerli numune formatındaysa ve daha önce eklenmediyse al
            if code not in seen_codes and len(code) > 5:
                # Kodun bulunduğu hücre indeksini belirle
                code_idx = -1
                for i, val in enumerate(row_cells):
                    if code in val:
                        code_idx = i
                        break

                tur = row_cells[code_idx + 1] if len(row_cells) > code_idx + 1 else "Beton / Sıva"
                yer = row_cells[code_idx + 2] if len(row_cells) > code_idx + 2 else "-"
                yontem = row_cells[code_idx + 3] if len(row_cells) > code_idx + 3 else "TS EN ISO 16000-7"
                strateji = row_cells[code_idx + 4] if len(row_cells) > code_idx + 4 else "Görsel ve Alansal"

                samples.append({
                    'kod': code,
                    'tur': tur,
                    'yer': yer,
                    'yontem': yontem,
                    'strateji': strateji
                })
                seen_codes.add(code)

    return info, samples

read_tutanak_details = parse_asbest_tutanak

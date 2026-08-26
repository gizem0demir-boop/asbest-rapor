import openpyxl
import re

def parse_asbest_tutanak(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    sheet = wb.active
    
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
    
    # Tüm sayfadaki satırları düz metin olarak topla
    for row in sheet.iter_rows(values_only=True):
        # Satırdaki None olmayan tüm değerleri al ve tek bir metin satırı yap
        row_cells = [str(cell).strip() for cell in row if cell is not None and str(cell).strip() != '']
        if not row_cells:
            continue
            
        full_row_text = " ".join(row_cells)
        
        # 1. MÜŞTERİ / FİRMA ADI
        if "Firma Adı" in full_row_text or "Müşteri" in full_row_text:
            match = re.search(r'(?:Firma|Müşteri)\s*Adı\s*:\s*([^:\n]+)', full_row_text, re.IGNORECASE)
            if match:
                val = match.group(1).split("Telefon")[0].split("Adres")[0].strip()
                if val: info['musteri_adi'] = val

        # 2. TEKLİF / TALEP NO (Örn: 26-08-5191 veya 26-08-519)
        if "Teklif" in full_row_text or "Talep" in full_row_text:
            match = re.search(r'\b(\d{2}-\d{2}-\d+)\b', full_row_text)
            if match:
                info['teklif_no'] = match.group(1)

        # 3. ADRES
        if "Adres" in full_row_text:
            match = re.search(r'Adres[i]?\s*:\s*([^:\n]+)', full_row_text, re.IGNORECASE)
            if match:
                val = match.group(1).split("Pafta")[0].split("Ada")[0].strip()
                if val and val != '-': info['adres'] = val

        # 4. PAFTA / ADA / PARSEL
        if "Pafta" in full_row_text or "Parsel" in full_row_text:
            p = re.search(r'Pafta\s*No\s*:\s*([^\s|]+)', full_row_text, re.IGNORECASE)
            a = re.search(r'Ada\s*No\s*:\s*([^\s|]+)', full_row_text, re.IGNORECASE)
            pr = re.search(r'Parsel\s*No\s*:\s*([^\s|]+)', full_row_text, re.IGNORECASE)
            if p: info['pafta'] = p.group(1).strip()
            if a: info['ada'] = a.group(1).strip()
            if pr: info['parsel'] = pr.group(1).strip()

        # 5. TARİH (GG.AA.YYYY)
        if "Tarih" in full_row_text:
            match = re.search(r'\b(\d{2}\.\d{2}\.\d{4})\b', full_row_text)
            if match:
                info['numune_tarihi'] = match.group(1)

        # 6. NUMUNE TABLOSU TARAMASI (NK. ile başlayan gerçek numune kodları)
        for i, cell_value in enumerate(row_cells):
            # Kod arama: NK.26.5038-01 gibi tam kalıpları yakalar
            code_match = re.search(r'\b(NK\.\d{2}\.\d+-\d+)\b', cell_value)
            if code_match:
                code = code_match.group(1)
                
                # Eğer aynı kod tekrar etmiyorsa listeye ekle
                if code not in seen_codes:
                    seen_codes.add(code)
                    
                    # Yan sütunlardaki bilgileri sırayla çek
                    tur = row_cells[i+1] if len(row_cells) > i+1 else "Beton / Sıva"
                    yer = row_cells[i+2] if len(row_cells) > i+2 else "-"
                    yontem = row_cells[i+3] if len(row_cells) > i+3 else "TS EN ISO 16000-7"
                    strateji = row_cells[i+4] if len(row_cells) > i+4 else "Görsel ve Alansal"
                    
                    # Eğer 'tur' alanı yanlışlıkla başka bir kod veya başlıksa temizle
                    if "NK." in tur or "Yöntem" in tur:
                        tur = "Beton / Sıva"

                    samples.append({
                        'kod': code,
                        'tur': tur,
                        'yer': yer,
                        'yontem': yontem,
                        'strateji': strateji
                    })

    return info, samples

# Modül bağlama takma adı
read_tutanak_details = parse_asbest_tutanak

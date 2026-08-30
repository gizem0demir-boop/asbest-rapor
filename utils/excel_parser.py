import pandas as pd
import re
import logging

def read_tutanak_details(tutanak_file):
    """Asbest Katı Numunesi Alma Tutanağı şablonundan bilgileri okur."""
    info = {"adres": "", "ada": "", "parsel": "", "musteri_adi": ""}
    try:
        if hasattr(tutanak_file, "seek"):
            tutanak_file.seek(0)
            
        df = pd.read_excel(tutanak_file, sheet_name="Table 1", header=None)
        
        for r_idx, row in df.iterrows():
            for c_idx, val in enumerate(row):
                if not isinstance(val, str):
                    continue
                val_str = val.strip()
                
                # Firma / Müşteri Adı
                if "Firma Adı:" in val_str:
                    parts = val_str.split("Firma Adı:")
                    if len(parts) > 1 and parts[1].strip():
                        info["musteri_adi"] = parts[1].strip()
                
                # Firma Adresi
                elif "Firma Adresi:" in val_str:
                    parts = val_str.split("Firma Adresi:")
                    if len(parts) > 1 and parts[1].strip():
                        info["adres"] = parts[1].strip()
                        
                # Ada No ve Parsel No taraması
                elif "Ada No" in val_str or "Ada:" in val_str:
                    # Aynı hücrede veya yan hücrede olabilir
                    match = re.search(r'(?:Ada\s*No[:\s]*)([0-9\w\-]+)', val_str, re.IGNORECASE)
                    if match:
                        info["ada"] = match.group(1)
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        info["ada"] = str(row.iloc[c_idx + 1]).strip()
                        
                elif "Parsel No" in val_str or "Parsel:" in val_str:
                    match = re.search(r'(?:Parsel\s*No[:\s]*)([0-9\w\-]+)', val_str, re.IGNORECASE)
                    if match:
                        info["parsel"] = match.group(1)
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        info["parsel"] = str(row.iloc[c_idx + 1]).strip()

        # Eğer yan hücrelere dağılmışsa genel satır taramasıyla da destekleyelim
        for r_idx, row in df.iterrows():
            row_text = " ".join([str(cell) for cell in row.values if pd.notna(cell)])
            
            if not info["musteri_adi"] and "Firma Adı" in row_text:
                m = re.search(r'Firma Adı[:\s]*([^\t\n]+?)(?=\s{2,}|Telefon|$)', row_text)
                if m:
                    info["musteri_adi"] = m.group(1).strip()
                    
            if not info["adres"] and "Firma Adresi" in row_text:
                m = re.search(r'Firma Adresi[:\s]*([^\t\n]+)', row_text)
                if m:
                    info["adres"] = m.group(1).strip()
                    
            if not info["parsel"] and "Parsel No" in row_text:
                m = re.search(r'Parsel\s*No[:\s]*([0-9\w\-]+)', row_text, re.IGNORECASE)
                if m:
                    info["parsel"] = m.group(1).strip()
                    
            if not info["ada"] and "Ada No" in row_text:
                m = re.search(r'Ada\s*No[:\s]*([0-9\w\-]+)', row_text, re.IGNORECASE)
                if m:
                    info["ada"] = m.group(1).strip()

        return info, df

    except Exception as e:
        logging.exception("Tutanak okunurken hata: %s", e)
        return {}, None

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
            row_text = " ".join([str(cell) for cell in row.values if pd.notna(cell)])
            
            # Firma Adı
            if "Firma Adı:" in row_text and not info["musteri_adi"]:
                m = re.search(r'Firma Adı[:\s]*([^\t\n]+?)(?=\s{2,}|Telefon|$)', row_text)
                if m:
                    info["musteri_adi"] = m.group(1).strip()
            
            # Firma Adresi
            if "Firma Adresi:" in row_text and not info["adres"]:
                m = re.search(r'Firma Adresi[:\s]*([^\t\n]+)', row_text)
                if m:
                    info["adres"] = m.group(1).strip()
                    
            # Ada ve Parsel Bilgileri
            if "Ada No" in row_text or "Parsel No" in row_text:
                for c_idx, val in enumerate(row):
                    if not isinstance(val, str):
                        continue
                    val_str = val.strip()
                    
                    if "Ada No" in val_str:
                        match = re.search(r'(?:Ada\s*No[:\s]*)([0-9\w\-]+)', val_str, re.IGNORECASE)
                        if match:
                            info["ada"] = match.group(1).strip()
                        elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                            info["ada"] = str(row.iloc[c_idx + 1]).strip()
                            
                    if "Parsel No" in val_str:
                        match = re.search(r'(?:Parsel\s*No[:\s]*)([0-9\w\-]+)', val_str, re.IGNORECASE)
                        if match:
                            info["parsel"] = match.group(1).strip()
                        elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                            info["parsel"] = str(row.iloc[c_idx + 1]).strip()

        return info, df

    except Exception as e:
        logging.exception("Tutanak okunurken hata: %s", e)
        return {}, None

def read_fenni_mesul_details(tutanak_file):
    """
    Fenni Mesul Taahhütnamesi sekmesi için yüklenen Excel dosyasından 
    adres, ada ve parsel bilgilerini özel olarak okur.
    """
    info = {
        "yapi_adresi": "-",
        "ada_parsel": "-"
    }
    
    try:
        if hasattr(tutanak_file, "seek"):
            tutanak_file.seek(0)
            
        df = pd.read_excel(tutanak_file, sheet_name="Table 1", header=None)
        
        ada_val = ""
        parsel_val = ""
        adres_val = ""
        
        for r_idx, row in df.iterrows():
            row_text = " ".join([str(cell) for cell in row.values if pd.notna(cell)])
            
            # Adres tespiti
            if "Firma Adresi:" in row_text or "Adresi:" in row_text:
                m = re.search(r'(?:Firma\s*Adresi|Adresi)[:\s]*([^\t\n]+)', row_text)
                if m:
                    adres_val = m.group(1).strip()
            
            # Hücre hücre Ada ve Parsel arama
            for c_idx, val in enumerate(row):
                if not isinstance(val, str):
                    continue
                val_str = val.strip()
                
                # Ada yakalama
                if "Ada No" in val_str or val_str.startswith("Ada"):
                    m_ada = re.search(r'(?:Ada\s*No[:\s]*)([0-9\w\-]+)', val_str, re.IGNORECASE)
                    if m_ada:
                        ada_val = m_ada.group(1).strip()
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        ada_val = str(row.iloc[c_idx + 1]).strip()
                        
                # Parsel yakalama
                if "Parsel No" in val_str or val_str.startswith("Parsel"):
                    m_parsel = re.search(r'(?:Parsel\s*No[:\s]*)([0-9\w\-]+)', val_str, re.IGNORECASE)
                    if m_parsel:
                        parsel_val = m_parsel.group(1).strip()
                    elif c_idx + 1 < len(row) and pd.notna(row.iloc[c_idx + 1]):
                        parsel_val = str(row.iloc[c_idx + 1]).strip()

        # Sonuçları yapılandır
        if adres_val:
            info["yapi_adresi"] = adres_val
            
        parts = []
        if ada_val and ada_val != "-":
            parts.append(f"Ada: {ada_val}")
        if parsel_val and parsel_val != "-":
            parts.append(f"Parsel: {parsel_val}")
            
        if parts:
            info["ada_parsel"] = " / ".join(parts)
        elif parsel_val:
            info["ada_parsel"] = parsel_val

        return info

    except Exception as e:
        logging.exception("Fenni Mesul tutanağı okunurken hata: %s", e)
        return {"yapi_adresi": "-", "ada_parsel": "-"}

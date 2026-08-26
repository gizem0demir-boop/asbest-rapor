import pandas as pd

def parse_asbest_tutanak(uploaded_file):
    """
    Excel tutanağından genel bilgileri ve numune listesini dinamik olarak tarar.
    """
    info = {
        "musteri_adi": "",
        "adres": "",
        "teklif_no": "",
        "numune_tarihi": "",
        "pafta": "",
        "ada": "",
        "parsel": ""
    }
    samples = []

    try:
        # Excel dosyasını başlık olmadan oku
        df = pd.read_excel(uploaded_file, header=None)
        
        # 1. Genel Bilgileri Excel İçinde Metin Arayarak Doldur
        for r in range(len(df)):
            for c in range(len(df.columns)):
                cell_val = str(df.iloc[r, c]).strip()
                
                if "MÜŞTERİ" in cell_val.upper() or "MAL SAHİBİ" in cell_val.upper():
                    if c + 1 < len(df.columns) and pd.notna(df.iloc[r, c + 1]):
                        info["musteri_adi"] = str(df.iloc[r, c + 1]).replace("nan", "").strip()
                
                elif "ADRES" in cell_val.upper():
                    if c + 1 < len(df.columns) and pd.notna(df.iloc[r, c + 1]):
                        info["adres"] = str(df.iloc[r, c + 1]).replace("nan", "").strip()
                        
                elif "TEKLİF NO" in cell_val.upper():
                    if c + 1 < len(df.columns) and pd.notna(df.iloc[r, c + 1]):
                        info["teklif_no"] = str(df.iloc[r, c + 1]).replace("nan", "").strip()
                        
                elif "TARİH" in cell_val.upper() and not info["numune_tarihi"]:
                    if c + 1 < len(df.columns) and pd.notna(df.iloc[r, c + 1]):
                        info["numune_tarihi"] = str(df.iloc[r, c + 1]).replace("nan", "").strip()

                elif "PAFTA" in cell_val.upper():
                    if c + 1 < len(df.columns) and pd.notna(df.iloc[r, c + 1]):
                        info["pafta"] = str(df.iloc[r, c + 1]).replace("nan", "").strip()
                        
                elif "ADA" in cell_val.upper():
                    if c + 1 < len(df.columns) and pd.notna(df.iloc[r, c + 1]):
                        info["ada"] = str(df.iloc[r, c + 1]).replace("nan", "").strip()
                        
                elif "PARSEL" in cell_val.upper():
                    if c + 1 < len(df.columns) and pd.notna(df.iloc[r, c + 1]):
                        info["parsel"] = str(df.iloc[r, c + 1]).replace("nan", "").strip()

        # 2. Numune Tablosunu Yakala
        # 'Numune Kodu' veya 'Numune No' geçen satırı tablo başlığı olarak bul
        header_row = None
        for r in range(len(df)):
            row_str = " ".join([str(val) for val in df.iloc[r].values]).upper()
            if "NUMUNE" in row_str and ("KOD" in row_str or "NO" in row_str or "TÜR" in row_str or "MALZEME" in row_str):
                header_row = r
                break

        if header_row is not None:
            # Başlıktan sonraki satırları numune olarak tara
            for r in range(header_row + 1, len(df)):
                row_vals = [str(val).replace("nan", "").strip() for val in df.iloc[r].values]
                # Satırın tamamen boş olmadığını kontrol et
                if any(row_vals):
                    kod = row_vals[0] if len(row_vals) > 0 else f"NUM-{r}"
                    tur = row_vals[1] if len(row_vals) > 1 else ""
                    yer = row_vals[2] if len(row_vals) > 2 else ""
                    
                    if kod: # Kod hücresi boş değilse listeye ekle
                        samples.append({
                            "kod": kod,
                            "tur": tur,
                            "yer": yer,
                            "yontem": "ISO 22262-1",
                            "strateji": "Rastgele Numune Alma"
                        })

    except Exception as e:
        print(f"Excel ayıklama hatası: {e}")

    return info, samples

import pandas as pd

def read_tutanak_details(uploaded_file):
    """
    Asbest Katı Numunesi Alma Tutanağı Excel dosyasını okuyup 
    şablona aktarılacak sözlük (dict) formatına çevirir.
    """
    try:
        # Excel'i oku (varsayılan ilk sayfa)
        df = pd.read_excel(uploaded_file)
        
        # Kodun ihtiyaç duyduğu alanları burada derliyoruz
        details = {
            "isveren": df.iloc[0, 1] if len(df) > 0 else "",
            "adres": df.iloc[1, 1] if len(df) > 1 else "",
            "numune_tarihi": df.iloc[2, 1] if len(df) > 2 else "",
            # İhtiyacına göre DataFrame hücre eşleşmelerini buraya ekleyebilirsin
        }
        return details
    except Exception as e:
        print(f"Excel okuma hatası: {e}")
        return {}
import pandas as pd

def read_tutanak_details(uploaded_file):
    """
    Asbest Katı Numunesi Alma Tutanağı Excel dosyasından 
    gerekli bilgileri okuyan fonksiyon.
    """
    try:
        df = pd.read_excel(uploaded_file)
        
        # Excel'den okunacak temel veriler
        data = {
            "bina_adi": str(df.iloc[0, 1]) if len(df) > 0 else "",
            "adres": str(df.iloc[1, 1]) if len(df) > 1 else "",
            "isveren": str(df.iloc[2, 1]) if len(df) > 2 else "",
        }
        return data
    except Exception as e:
        print(f"Excel okunurken hata oluştu: {e}")
        return {}

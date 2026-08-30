def genisletilmis_tutanak_oku(tutanak_file):
    """Yüklenen Excel tabanlı tutanak dosyasını okur ve ilgili alanları döndürür."""
    if tutanak_file is None:
        return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

    try:
        if hasattr(tutanak_file, "seek"):
            tutanak_file.seek(0)

        if read_tutanak_details is None:
            logging.warning("read_tutanak_details fonksiyonu yüklenememiş.")
            return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

        res = read_tutanak_details(tutanak_file)
        
        info_dict = {}
        if isinstance(res, tuple) and len(res) > 0:
            info_dict = res[0] if isinstance(res[0], dict) else {}
        elif isinstance(res, dict):
            info_dict = res

        if info_dict:
            ada_val = info_dict.get("ada") or info_dict.get("Ada") or "-"
            parsel_val = info_dict.get("parsel") or info_dict.get("Parsel") or "-"
            ada_parsel_str = f"{ada_val} Ada {parsel_val} Parsel" if ada_val != "-" or parsel_val != "-" else "-"
            
            return {
                "yapi_adresi": info_dict.get("adres") or info_dict.get("Adres") or "",
                "ada_parsel": ada_parsel_str,
                "musteri_adi": info_dict.get("musteri_adi") or info_dict.get("Musteri_Adi") or ""
            }

    except Exception as e:
        logging.exception("Excel tutanak dosyası okunurken hata oluştu: %s", e)

    return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

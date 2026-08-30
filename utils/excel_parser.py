import logging

def genisletilmis_tutanak_oku(tutanak_file):
    """Yüklenen dosyanın türüne göre doğru ayrıştırıcıyı seçer, 
    imleci başa alır ve hataları yakalayarak her durumda veri döndürür.
    """
    if tutanak_file is None:
        return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

    try:
        # Streamlit file-like nesnelerinin imlecini mutlaka başa alıyoruz
        if hasattr(tutanak_file, "seek"):
            tutanak_file.seek(0)

        file_name = getattr(tutanak_file, "name", "").lower()

        # 1. PDF Dosyaları İçin
        if file_name.endswith(".pdf"):
            if parse_asbestos_pdf_report is None:
                logging.warning("parse_asbestos_pdf_report fonksiyonu yüklenememiş.")
                return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

            temp_path = "temp_yikim_parse.pdf"
            with open(temp_path, "wb") as f:
                f.write(tutanak_file.getbuffer())
            
            try:
                pdf_data = parse_asbestos_pdf_report(temp_path)
                if not pdf_data or not isinstance(pdf_data, dict):
                    return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}
                
                ada_val = pdf_data.get("ada") or pdf_data.get("Ada") or "-"
                parsel_val = pdf_data.get("parsel") or pdf_data.get("Parsel") or "-"
                ada_parsel_str = f"{ada_val} Ada {parsel_val} Parsel" if ada_val != "-" or parsel_val != "-" else "-"
                
                return {
                    "yapi_adresi": pdf_data.get("adres") or pdf_data.get("Adres") or "",
                    "ada_parsel": ada_parsel_str,
                    "musteri_adi": pdf_data.get("musteri_adi") or pdf_data.get("Musteri_Adi") or "",
                }
            except Exception as e:
                logging.exception("PDF parse edilirken hata: %s", e)
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)

        # 2. Excel / Word / Diğer Dosyalar İçin
        else:
            if read_tutanak_details is None:
                logging.warning("read_tutanak_details fonksiyonu yüklenememiş.")
                return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

            try:
                res = read_tutanak_details(tutanak_file)
                
                # Fonksiyondan dönen yapının formatına göre esnek çözüm
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
                logging.exception("Tutanak detayları okunurken hata: %s", e)

    except Exception as general_e:
        logging.exception("genisletilmis_tutanak_oku genel hata: %s", general_e)

    return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

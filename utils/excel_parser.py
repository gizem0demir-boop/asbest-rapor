import tempfile
from io import BytesIO

def genisletilmis_tutanak_oku(tutanak_file):
    """UploadedFile'ı güvenli şekilde işleyip parse fonksiyonuna temp dosya yolunu verir.
    Hem PDF, hem Excel (.xls/.xlsx) hem de diğer formatlar için daha kararlı çalışır.
    """
    if tutanak_file is None:
        return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

    # Debug: tip ve isim
    try:
        st.write("DEBUG: uploaded file type:", type(tutanak_file))
        st.write("DEBUG: uploaded file name:", getattr(tutanak_file, "name", None))
    except Exception:
        pass

    name = getattr(tutanak_file, "name", "").lower()
    try:
        # PDF için mevcut yaklaşımla temp dosya oluştur
        if name.endswith(".pdf"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(tutanak_file.getbuffer())
                tmp_path = tmp.name

            try:
                pdf_data = parse_asbestos_pdf_report(tmp_path)
                if not pdf_data or not isinstance(pdf_data, dict):
                    return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}
                ada_val = pdf_data.get("ada", "-")
                parsel_val = pdf_data.get("parsel", "-")
                ada_parsel_str = f"{ada_val} Ada {parsel_val} Parsel" if ada_val != "-" or parsel_val != "-" else "-"
                return {
                    "yapi_adresi": pdf_data.get("adres", "-"),
                    "ada_parsel": ada_parsel_str,
                    "musteri_adi": pdf_data.get("musteri_adi", ""),
                }
            finally:
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

        # Excel dosyaları için temp path oluşturup parse fonksiyonuna ver
        elif name.endswith((".xlsx", ".xls")):
            suffix = os.path.splitext(name)[1] or ".xlsx"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(tutanak_file.getbuffer())
                tmp_path = tmp.name

            try:
                # read_tutanak_details parse fonksiyonunuz path veya file-like kabul ediyorsa path ile tutarlı çalışır
                res = read_tutanak_details(tmp_path)
                if isinstance(res, tuple) and len(res) >= 1:
                    info_dict = res[0]
                elif isinstance(res, dict):
                    info_dict = res
                else:
                    st.write("DEBUG: read_tutanak_details döndüğü değer beklenmedik:", res)
                    return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

                ada_val = info_dict.get("ada", "-")
                parsel_val = info_dict.get("parsel", "-")
                ada_parsel_str = f"{ada_val} Ada {parsel_val} Parsel" if ada_val != "-" or parsel_val != "-" else "-"
                return {
                    "yapi_adresi": info_dict.get("adres", ""),
                    "ada_parsel": ada_parsel_str,
                    "musteri_adi": info_dict.get("musteri_adi", "")
                }
            except Exception as e:
                st.write("DEBUG: Excel parse hatası:", e)
            finally:
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

        else:
            # Diğer formatlarda önce file-like dene, hata olursa temp file ile tekrar dene
            try:
                res = read_tutanak_details(tutanak_file)
                if isinstance(res, tuple) and len(res) >= 1:
                    info_dict = res[0]
                elif isinstance(res, dict):
                    info_dict = res
                else:
                    st.write("DEBUG: read_tutanak_details (other) beklenmedik çıktı:", res)
                    return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

                ada_val = info_dict.get("ada", "-")
                parsel_val = info_dict.get("parsel", "-")
                ada_parsel_str = f"{ada_val} Ada {parsel_val} Parsel" if ada_val != "-" or parsel_val != "-" else "-"
                return {
                    "yapi_adresi": info_dict.get("adres", ""),
                    "ada_parsel": ada_parsel_str,
                    "musteri_adi": info_dict.get("musteri_adi", "")
                }
            except Exception as e:
                st.write("DEBUG: read_tutanak_details (other) hata:", e)

    except Exception as e:
        st.write("DEBUG: genisletilmis_tutanak_oku hata:", e)

    return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

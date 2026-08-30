import datetime
import os
import re
import sys
import tempfile
import logging

import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

# ensure project root is on path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Safe imports from utils: log failures and keep None so app doesn't crash on import
try:
    from utils.excel_parser import read_tutanak_details
except Exception as e:
    logging.exception("Failed to import read_tutanak_details: %s", e)
    read_tutanak_details = None

try:
    from utils.pdf_parser import parse_asbestos_pdf_report
except Exception as e:
    logging.exception("Failed to import parse_asbestos_pdf_report: %s", e)
    parse_asbestos_pdf_report = None

EXCEL_VT_YOLU = "veritabani.xlsx"

SUPPORTED_FILE_TYPES = [
    "xlsx",
    "xls",
    "docx",
    "doc",
    "pdf",
    "jpg",
    "jpeg",
    "png",
]


def sayiyi_yaziya_cevir(tutar_str):
    try:
        rakamlar = re.findall(r"\d+", str(tutar_str))
        if not rakamlar:
            return tutar_str

        tutar = int("".join(rakamlar))
        if tutar == 0:
            return "Sıfır Türk Lirası"

        birler = ["", "Bir", "İki", "Üç", "Dört", "Beş", "Altı", "Yedi", "Sekiz", "Dokuz"]
        onlar = ["", "On", "Yirmi", "Otuz", "Kırk", "Elli", "Altmış", "Yetmiş", "Sekiz", "Doksan"]

        def yuz_str(x):
            y = x // 100
            o = (x % 100) // 10
            b = x % 10
            s = ""
            if y > 0:
                s += ("BirYüz" if y == 1 else birler[y] + "Yüz")
            if o > 0:
                s += onlar[o]
            if b > 0:
                s += birler[b]
            return s

        milyon = (tutar // 1_000_000) % 1000
        bin_grubu = (tutar // 1_000) % 1000
        kalan = tutar % 1000

        parcalar = []
        if milyon > 0:
            parcalar.append("BirMilyon" if milyon == 1 else yuz_str(milyon) + "Milyon")
        if bin_grubu > 0:
            parcalar.append("Bin" if bin_grubu == 1 else yuz_str(bin_grubu) + "Bin")
        if kalan > 0:
            parcalar.append(yuz_str(kalan))

        tam_metin = "".join(parcalar)
        bosluklu = re.sub(r"(?<!^)(?=[A-Z])", " ", tam_metin)
        return bosluklu + " Türk Lirası"
    except Exception:
        return tutar_str


def genisletilmis_tutanak_oku(tutanak_file):
    """
    Return dict: { yapi_adresi, ada_parsel, musteri_adi }.
    Uses utils parsers when available; writes a temp file for robust parsing.
    """
    if tutanak_file is None:
        return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}

    name = getattr(tutanak_file, "name", "") or ""
    name = name.lower()

    # PDF
    if name.endswith(".pdf"):
        if parse_asbestos_pdf_report is None:
            logging.error("PDF parser not available")
            return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}
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
            return {"yapi_adresi": pdf_data.get("adres", "-"), "ada_parsel": ada_parsel_str, "musteri_adi": pdf_data.get("musteri_adi", "")}
        except Exception as e:
            logging.exception("PDF parse error: %s", e)
            return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}
        finally:
            try:
                os.remove(tmp_path)
            except Exception:
                pass

    # Excel
    if name.endswith((".xlsx", ".xls")):
        if read_tutanak_details is None:
            logging.error("Excel parser not available")
            return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}
        suffix = os.path.splitext(name)[1] or ".xlsx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(tutanak_file.getbuffer())
            tmp_path = tmp.name
        try:
            res = read_tutanak_details(tmp_path)
            if isinstance(res, tuple) and len(res) >= 1:
                info_dict = res[0]
            elif isinstance(res, dict):
                info_dict = res
            else:
                logging.warning("read_tutanak_details returned unexpected type: %s", type(res))
                return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}
            ada_val = info_dict.get("ada", "-")
            parsel_val = info_dict.get("parsel", "-")
            ada_parsel_str = f"{ada_val} Ada {parsel_val} Parsel" if ada_val != "-" or parsel_val != "-" else "-"
            return {"yapi_adresi": info_dict.get("adres", ""), "ada_parsel": ada_parsel_str, "musteri_adi": info_dict.get("musteri_adi", "")}
        except Exception as e:
            logging.exception("Excel parse error: %s", e)
            return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}
        finally:
            try:
                os.remove(tmp_path)
            except Exception:
                pass

    # Fallback: try parser with file-like if available
    try:
        if read_tutanak_details is not None:
            res = read_tutanak_details(tutanak_file)
            if isinstance(res, tuple) and len(res) >= 1:
                info_dict = res[0]
            elif isinstance(res, dict):
                info_dict = res
            else:
                return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}
            ada_val = info_dict.get("ada", "-")
            parsel_val = info_dict.get("parsel", "-")
            ada_parsel_str = f"{ada_val} Ada {parsel_val} Parsel" if ada_val != "-" or parsel_val != "-" else "-"
            return {"yapi_adresi": info_dict.get("adres", ""), "ada_parsel": ada_parsel_str, "musteri_adi": info_dict.get("musteri_adi", "")}
    except Exception as e:
        logging.exception("Fallback parser error: %s", e)

    return {"yapi_adresi": "", "ada_parsel": "", "musteri_adi": ""}


@st.cache_data(ttl=60)
def veritabani_yukle():
    if not os.path.exists(EXCEL_VT_YOLU):
        st.error(f"❌ '{EXCEL_VT_YOLU}' bulunamadı. Lütfen repo dizinine Excel dosyasını ekleyin.")
        return pd.DataFrame(), pd.DataFrame()
    try:
        df_muellif = pd.read_excel(EXCEL_VT_YOLU, sheet_name=0)
        df_muteahhit = pd.read_excel(EXCEL_VT_YOLU, sheet_name=1)
        return df_muellif, df_mute


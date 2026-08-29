from collections import OrderedDict
import io
import os
import re
from docxtpl import InlineImage
from docx.shared import Mm
import pandas as pd
from PIL import Image, ImageOps

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def _get_file_extension(file_or_path):
    if file_or_path is None:
        return ""
    file_name = getattr(file_or_path, "name", str(file_or_path))
    return os.path.splitext(file_name)[1].lower()


# --- EXCEL FORMAT HATALARINI ÖNLEYEN GÜVENLİ OKUYUCU ---
def safe_read_excel(file_or_path, sheet_name=0, **kwargs):
    ext = _get_file_extension(file_or_path)

    # PDF veya excel dışı bir dosya geldiyse pandas'ı asla çalıştırma
    if ext in [".pdf", ".docx", ".doc", ".jpg", ".jpeg", ".png"]:
        raise ValueError(
            "Yüklenen dosya Excel formatında değil. Lütfen geçerli bir Excel"
            " dosyası (.xlsx / .xls) yükleyin."
        )

    if hasattr(file_or_path, "seek"):
        file_or_path.seek(0)

    file_name = getattr(file_or_path, "name", str(file_or_path)).lower()

    if file_name.endswith(".xlsx"):
        try:
            return pd.read_excel(
                file_or_path, sheet_name=sheet_name, engine="openpyxl", **kwargs
            )
        except Exception:
            if hasattr(file_or_path, "seek"):
                file_or_path.seek(0)
            return pd.read_excel(
                file_or_path, sheet_name=sheet_name, engine="xlrd", **kwargs
            )

    elif file_name.endswith(".xls"):
        try:
            return pd.read_excel(
                file_or_path, sheet_name=sheet_name, engine="xlrd", **kwargs
            )
        except Exception:
            if hasattr(file_or_path, "seek"):
                file_or_path.seek(0)
            return pd.read_excel(
                file_or_path, sheet_name=sheet_name, engine="openpyxl", **kwargs
            )
    else:
        try:
            return pd.read_excel(
                file_or_path, sheet_name=sheet_name, engine="openpyxl", **kwargs
            )
        except Exception:
            if hasattr(file_or_path, "seek"):
                file_or_path.seek(0)
            return pd.read_excel(
                file_or_path, sheet_name=sheet_name, engine="xlrd", **kwargs
            )


# --------------------------------------------------------


def process_and_get_image(doc, uploaded_file, width_cm=6.5, height_cm=5.0):
    if uploaded_file is None:
        return ""
    try:
        img = Image.open(uploaded_file)
        img = ImageOps.exif_transpose(img)
        img.thumbnail((1200, 1200))

        img_byte_arr = io.BytesIO()
        img_format = img.format if img.format else "JPEG"
        img.save(img_byte_arr, format=img_format, quality=85)
        img_byte_arr.seek(0)

        return InlineImage(
            doc, img_byte_arr, width=Mm(width_cm * 10), height=Mm(height_cm * 10)
        )
    except Exception:
        return ""


def generate_bolum_summary(samples):
    place_counts = OrderedDict()
    for s in samples:
        yer = s["yer"] if s["yer"] and s["yer"] != "-" else "Belirtilmedi"
        place_counts[yer] = place_counts.get(yer, 0) + 1
    return [{"yer": yer, "sayi": sayi} for yer, sayi in place_counts.items()]


def read_tutanak_details(tutanak_path):
    ext = _get_file_extension(tutanak_path)
    if ext in [".pdf", ".docx", ".doc", ".jpg", ".jpeg", ".png"]:
        return {
            "musteri_adi": "",
            "MUSTERI_ADI": "",
            "firma_adi": "",
            "FIRMA_ADI": "",
            "adres": "",
            "ADRES": "",
            "santiye_adresi": "",
            "SANTIYE_ADRESI": "",
            "pafta": "-",
            "ada": "-",
            "parsel": "-",
            "pafta_ada_parsel": "- / - / -",
            "PAFTA_ADA_PARSEL": "- / - / -",
        }

    try:
        df = safe_read_excel(tutanak_path, sheet_name="Table 1", header=None)
    except Exception:
        xls = pd.ExcelFile(tutanak_path)
        df = safe_read_excel(xls, sheet_name=xls.sheet_names[0], header=None)

    raw_firma = str(df.iloc[4, 0]) if pd.notna(df.iloc[4, 0]) else ""
    musteri_adi = raw_firma.replace("Firma Adı:", "").strip()

    raw_adres = str(df.iloc[5, 0]) if pd.notna(df.iloc[5, 0]) else ""
    adres = raw_adres.replace("Firma Adresi:", "").strip()

    raw_pafta = str(df.iloc[6, 0]) if pd.notna(df.iloc[6, 0]) else "-"
    raw_ada = str(df.iloc[6, 4]) if pd.notna(df.iloc[6, 4]) else "-"
    raw_parsel = str(df.iloc[6, 8]) if pd.notna(df.iloc[6, 8]) else "-"

    pafta = raw_pafta.replace("Pafta No:", "").strip() or "-"
    ada = raw_ada.replace("Ada No:", "").strip() or "-"
    parsel = raw_parsel.replace("Parsel No:", "").strip() or "-"
    pafta_ada_parsel = f"{pafta} / {ada} / {parsel}"

    return {
        "musteri_adi": musteri_adi,
        "MUSTERI_ADI": musteri_adi,
        "firma_adi": musteri_adi,
        "FIRMA_ADI": musteri_adi,
        "adres": adres,
        "ADRES": adres,
        "santiye_adresi": adres,
        "SANTIYE_ADRESI": adres,
        "pafta": pafta,
        "ada": ada,
        "parsel": parsel,
        "pafta_ada_parsel": pafta_ada_parsel,
        "PAFTA_ADA_PARSEL": pafta_ada_parsel,
    }


def parse_asbest_tutanak(file):
    ext = _get_file_extension(file)
    if ext in [".pdf", ".docx", ".doc", ".jpg", ".jpeg", ".png"]:
        return {
            "musteri_adi": "-",
            "adres": "-",
            "pafta": "-",
            "ada": "-",
            "parsel": "-",
            "numune_tarihi": "-",
            "teklif_no": "-",
            "telefon": "-",
        }, []

    df_raw = safe_read_excel(file, header=None)

    info = {
        "musteri_adi": "ABC İnşaat",
        "adres": "-",
        "pafta": "-",
        "ada": "-",
        "parsel": "-",
        "numune_tarihi": "20.08.2026",
        "teklif_no": "26-08-5191",
        "telefon": "-",
    }

    for idx in range(min(10, len(df_raw))):
        row_values = [str(x) for x in df_raw.iloc[idx].values if pd.notna(x)]
        row_text = " ".join(row_values)

        if "Talep Numarası" in row_text and idx + 1 < len(df_raw):
            val = str(df_raw.iloc[idx + 1].values[0])
            if val and val != "nan":
                info["teklif_no"] = val.strip()

        if "Firma Adı:" in row_text:
            m = re.search(r"Firma Adı:\s*(.*?)(?:Telefon|$)", row_text)
            if m and m.group(1).strip():
                info["musteri_adi"] = m.group(1).strip()

        if "Telefon Numarası:" in row_text:
            m = re.search(r"Telefon Numarası:\s*(.*)", row_text)
            if m and m.group(1).strip():
                info["telefon"] = m.group(1).strip()

        if "Firma Adresi:" in row_text:
            m = re.search(r"Firma Adresi:\s*(.*)", row_text)
            if m and m.group(1).strip():
                info["adres"] = m.group(1).strip()

        if "Pafta No:" in row_text or "Parsel No:" in row_text:
            p = re.search(
                r"Pafta\s*No:\s*([^\s|]*)(?=\s*Ada|$)", row_text, re.IGNORECASE
            )
            a = re.search(
                r"Ada\s*No:\s*([^\s|]*)(?=\s*Parsel|$)", row_text, re.IGNORECASE
            )
            pr = re.search(
                r"Parsel\s*No:\s*([^\s|]*)(?=$)", row_text, re.IGNORECASE
            )
            if p and p.group(1).strip():
                info["pafta"] = p.group(1).strip()
            if a and a.group(1).strip():
                info["ada"] = a.group(1).strip()
            if pr and pr.group(1).strip():
                info["parsel"] = pr.group(1).strip()

        if "Tarih" in row_text:
            for cell in row_values:
                if re.match(r"\d{2}\.\d{2}\.\d{4}", cell):
                    info["numune_tarihi"] = cell

    samples = []
    for idx in range(len(df_raw)):
        row = df_raw.iloc[idx]
        row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
        code_match = re.search(r"NK\.\d+\.\d+-\d+", row_str)

        if code_match:
            code = code_match.group(0)
            non_empty = [
                str(x).strip()
                for x in row.values
                if pd.notna(x) and str(x).strip() != ""
            ]
            if len(non_empty) >= 3 and any(
                k in non_empty[1] for k in ["NK.", "NK"]
            ):
                samples.append({
                    "kod": code,
                    "tur": (
                        non_empty[2] if len(non_empty) > 2 else "Beton / Sıva"
                    ),
                    "yer": non_empty[3] if len(non_empty) > 3 else "-",
                    "yontem": non_empty[4] if len(non_empty) > 4 else "-",
                    "strateji": non_empty[5] if len(non_empty) > 5 else "-",
                })

    return info, samples

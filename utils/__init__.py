import os
import logging

# Mutlak ve göreceli importu destekleyen güvenli yapı
try:
    from utils.excel_parser import parse_asbest_tutanak, read_tutanak_details
except Exception as e:
    try:
        from .excel_parser import parse_asbest_tutanak, read_tutanak_details
    except Exception as inner_e:
        logging.exception("utils/__init__ import error: %s / %s", e, inner_e)
        parse_asbest_tutanak = None
        read_tutanak_details = None

# Upload klasörü
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def generate_bolum_summary(samples):
    from collections import OrderedDict
    place_counts = OrderedDict()
    for s in samples:
        yer = s.get('yer') if s.get('yer') and s.get('yer') != '-' else 'Belirtilmedi'
        place_counts[yer] = place_counts.get(yer, 0) + 1

    bolum_summary = []
    for yer, sayi in place_counts.items():
        bolum_summary.append({
            'yer': yer,
            'sayi': sayi
        })
    return bolum_summary


def process_and_get_image(doc, uploaded_file, width_cm=6.5, height_cm=5.0):
    """
    DocxTemplate için InlineImage döndüren yardımcı fonksiyon.
    uploaded_file bir Streamlit UploadedFile veya file-like olabilir.
    Başarısızlıkta boş string döner.
    """
    if uploaded_file is None:
        return ""
    try:
        import io
        from PIL import Image, ImageOps
        from docxtpl import InlineImage
        from docx.shared import Mm

        # Eğer uploaded_file bir path ise direkt aç, değilse file-like olarak aç
        if isinstance(uploaded_file, str) and os.path.exists(uploaded_file):
            img = Image.open(uploaded_file)
        else:
            # Streamlit UploadedFile gibi file-like objeler için imleci başa al
            uploaded_file.seek(0)
            img = Image.open(uploaded_file)

        img = ImageOps.exif_transpose(img)
        img.thumbnail((1200, 1200))

        img_byte_arr = io.BytesIO()
        img_format = img.format if img.format else 'JPEG'
        img.save(img_byte_arr, format=img_format, quality=85)
        img_byte_arr.seek(0)

        return InlineImage(
            doc,
            img_byte_arr,
            width=Mm(width_cm * 10),
            height=Mm(height_cm * 10)
        )
    except Exception as e:
        logging.exception("Görsel işleme hatası: %s", e)
        return ""

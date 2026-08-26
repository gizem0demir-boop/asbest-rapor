# utils/__init__.py
import os
from .excel_parser import parse_asbest_tutanak, read_tutanak_details

def generate_bolum_summary(samples):
    return f"Toplam {len(samples)} adet numune analize tabi tutulmuştur."

def process_and_get_image(tpl, uploaded_file, width_cm=6.0, height_cm=5.0):
    if uploaded_file is None:
        return ""
    try:
        from docxtpl import InlineImage
        from docx.shared import Cm
        return InlineImage(tpl, uploaded_file, width=Cm(width_cm), height=Cm(height_cm))
    except Exception as e:
        print(f"Görsel işleme hatası: {e}")
        return ""

# utils/__init__.py
import os
from .excel_parser import parse_asbest_tutanak, read_tutanak_details

# Toz ve diğer modüllerin aradığı yükleme klasörü tanımı
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def generate_bolum_summary(samples):
    return f"Toplam {len(samples)} adet numune analize tabi tutulmuştur."

def process_and_get_image(tpl, uploaded_file, width_cm=6.0, height_cm=5.0):
    if uploaded_file is None:
        return ""
    try:
        from docxtpl import InlineImage
        from docx.shared import Cm
        return InlineImage(tpl, uploaded_file, width=Cm(width_cm), height=Cm(height_cm))
    except Exception as e:
        print(f"Görsel işleme hatası: {e}")
        return ""

import os
from .excel_parser import parse_asbest_tutanak, read_tutanak_details

UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def generate_bolum_summary(samples):
    from collections import OrderedDict
    place_counts = OrderedDict()
    for s in samples:
        yer = s['yer'] if s.get('yer') and s['yer'] != '-' else 'Belirtilmedi'
        place_counts[yer] = place_counts.get(yer, 0) + 1
    
    bolum_summary = []
    for yer, sayi in place_counts.items():
        bolum_summary.append({
            'yer': yer,
            'sayi': sayi
        })
    return bolum_summary

def process_and_get_image(doc, uploaded_file, width_cm=6.5, height_cm=5.0):
    if uploaded_file is None:
        return ""
    try:
        import io
        from PIL import Image, ImageOps
        from docxtpl import InlineImage
        from docx.shared import Mm

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
        print(f"Görsel işleme hatası: {e}")
        return ""

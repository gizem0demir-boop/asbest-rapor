# utils/__init__.py
import os
from .excel_parser import read_tutanak_details

UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")

def parse_asbest_tutanak(uploaded_file):
    # excel_parser.py içindeki gerçek okuma fonksiyonuna yönlendirir
    return read_tutanak_details(uploaded_file)

def generate_bolum_summary(*args, **kwargs):
    return ""

def process_and_get_image(*args, **kwargs):
    return None

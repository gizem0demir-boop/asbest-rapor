# utils/__init__.py
import os
from .excel_parser import read_tutanak_details

UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")

def parse_asbest_tutanak(uploaded_file):
    # excel_parser'dan dönen veriyi alır
    res = read_tutanak_details(uploaded_file)
    if isinstance(res, tuple) and len(res) == 2:
        return res
    elif isinstance(res, dict):
        return res, res.get("samples", [])
    return {}, []

def generate_bolum_summary(*args, **kwargs):
    return ""

def process_and_get_image(*args, **kwargs):
    return None

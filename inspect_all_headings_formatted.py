import os
import sys
from docx import Document

sys.stdout.reconfigure(encoding='utf-8')

file_path = r"c:\Users\The Linh\Desktop\2022604088_NhuDinhChien_BC_DATN_formatted_test.docx"
doc = Document(file_path)

print("--- DANH SÁCH HEADING TRONG FILE FORMATTED TRÊN DESKTOP ---")
for idx, p in enumerate(doc.paragraphs):
    style_name = p.style.name if p.style else 'Normal'
    if style_name.startswith('Heading'):
        print(f"[{idx:03d}][Style: {style_name}]: {repr(p.text)}")

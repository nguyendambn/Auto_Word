import os
import sys
import re
from docx import Document

sys.stdout.reconfigure(encoding='utf-8')

input_path = r"c:\Users\The Linh\Desktop\2022604488_MaiXuanBac.docx"
doc = Document(input_path)

print("--- TÌM PARAGRAPH CHỨA '1.1.1.3' HOẶC 'Computer Vision' ---")
for idx, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if "1.1.1.3" in text or "Computer Vision" in text:
        print(f"[{idx:03d}][Style: {p.style.name}]: {repr(text)}")
        # Check XML properties if possible
        pPr = p._p.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
        if pPr is not None:
            outlineLvl = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}outlineLvl')
            val = outlineLvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if outlineLvl is not None else 'None'
            print(f"  outlineLvl: {val}")

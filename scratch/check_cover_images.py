import sys
import os

sys.path.insert(0, r'C:\Users\The Linh\Desktop\Auto-Word')
import docx
from docx_processor import paragraph_has_image

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\The Linh\Desktop\Auto-Word\_verify_v2.docx')

for idx in range(28):
    p = doc.paragraphs[idx]
    has_img = paragraph_has_image(p)
    print(f"Index {idx:2d} | Has Image: {has_img:5s} | Text: {repr(p.text)}")

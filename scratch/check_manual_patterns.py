import sys
import docx
import re

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\The Linh\Desktop\CNTT_2022607319_NguyenVanDam_Baocao.docx')

print("--- REGEX MATCHES IN BODY PARAGRAPHS ---")
patterns = [
    r'^\s*-\s*\d+\s*-\s*$',
    r'^\s*trang\s*\d+\s*$',
    r'^\s*\d+\s*/\s*\d+\s*$',
    r'^\s*trang\s*\d+\s*/\s*\d+\s*$'
]
for idx, p in enumerate(doc.paragraphs):
    text = p.text.strip().lower()
    if not text:
        continue
    for pat in patterns:
        if re.match(pat, text):
            print(f"Index {idx:4d} | Pattern: {pat} | Text: {p.text}")

import sys
import docx
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\The Linh\Desktop\CNTT_2022607319_NguyenVanDam_Baocao.docx')

print("--- PARAGRAPHS WITH OUTLINE LEVEL OR HEADING STYLE ---")
for idx, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    pPr = p._p.find(qn('w:pPr'))
    outline_val = None
    if pPr is not None:
        outlineLvl = pPr.find(qn('w:outlineLvl'))
        if outlineLvl is not None:
            outline_val = outlineLvl.get(qn('w:val'))
            
    style_name = p.style.name if p.style else ''
    
    if outline_val is not None or style_name.startswith("Heading") or text.startswith("CHƯƠNG") or text.startswith("MỞ ĐẦU"):
        print(f"Index {idx:4d} | Style: {style_name:20s} | Outline: {outline_val} | Text: {text[:60]}")

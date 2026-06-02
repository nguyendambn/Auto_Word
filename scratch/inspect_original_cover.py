import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\The Linh\Desktop\CNTT_2022607319_NguyenVanDam_Baocao.docx')

print("--- ORIGINAL COVER PARAGRAPHS ---")
for idx, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    pPr = p._p.find(docx.oxml.ns.qn('w:pPr'))
    has_sect = pPr is not None and pPr.find(docx.oxml.ns.qn('w:sectPr')) is not None
    
    print(f"Index {idx:2d} | Text: {repr(p.text)}")
    if has_sect:
        print("--- END OF SECTION 1 ---")
        break

import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\The Linh\Desktop\Auto-Word\_verify_headings.docx')

# Let's find all paragraphs in the first section (cover section)
first_section = doc.sections[0]
print("--- COVER PAGE PARAGRAPHS ---")
count = 0
for idx, p in enumerate(doc.paragraphs):
    # Check if there is a section break
    pPr = p._p.find(docx.oxml.ns.qn('w:pPr'))
    has_sect = pPr is not None and pPr.find(docx.oxml.ns.qn('w:sectPr')) is not None
    
    print(f"Index {idx:3d} | Style: {p.style.name:20s} | Text: {repr(p.text)}")
    count += 1
    if has_sect:
        print(f"--- END OF SECTION 1 (at index {idx}) ---")
        break

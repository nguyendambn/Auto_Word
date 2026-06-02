import sys
import os
import docx
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, r'C:\Users\The Linh\Desktop\Auto-Word')
from docx_processor import format_document

INPUT = r'C:\Users\The Linh\Desktop\CNTT_2022607319_NguyenVanDam_Baocao.docx'
OUTPUT = r'C:\Users\The Linh\Desktop\Auto-Word\_verify_headings.docx'

opts = {
    'font_name': 'Times New Roman', 'body_size': 13, 'line_spacing': 1.3,
    'space_before': 0, 'space_after': 6, 'first_line_indent': 10,
    'margin_top': 20, 'margin_bottom': 20, 'margin_left': 30, 'margin_right': 15,
    'heading1_size': 14, 'heading1_bold': True, 'heading1_uppercase': True,
    'heading2_size': 14, 'heading2_bold': True,
    'heading3_size': 13, 'heading3_italic': False,
    'auto_number_headings': True,  # Keep this True to test autonumbering!
    'format_admin_parts': False,
    'add_page_numbers': True, 'alignment': 'justify',
}

print("=== Formatting document with auto_number_headings=True ===")
format_document(INPUT, OUTPUT, opts)
print("=== Done formatting ===\n")

doc = docx.Document(OUTPUT)
print("=== HEADINGS IN FORMATTED DOCUMENT ===")
for idx, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if not text:
        continue
    style = p.style.name if p.style else ''
    if style.startswith("Heading"):
        print(f"Index {idx:4d} | Style: {style:15s} | Text: {text}")

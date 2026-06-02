import sys
import docx
import re

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\The Linh\Desktop\Auto-Word\_verify_headings.docx')

print("--- INDENTATION OF LIST ITEMS IN FORMATTED DOCUMENT ---")
for idx, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if not text:
        continue
    # Let's inspect paragraphs that start with numbers (e.g. 1. Tên đề tài, 2. Lý do chọn đề tài)
    if re.match(r'^\d+\.', text):
        indent = p.paragraph_format.first_line_indent
        indent_str = f"{indent.mm:.1f} mm" if indent is not None else "None"
        print(f"Index {idx:4d} | Indent: {indent_str:10s} | Text: {text}")

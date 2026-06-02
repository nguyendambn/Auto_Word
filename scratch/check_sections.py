import docx
import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, r'C:\Users\The Linh\Desktop\Auto-Word')

from docx_processor import partition_paragraphs_by_section, classify_sections

doc = docx.Document(r'C:\Users\The Linh\Desktop\Auto-Word\_verify_v2.docx')
parts = partition_paragraphs_by_section(doc)
sect_types = classify_sections(doc)

print("=== SECTIONS SUMMARY ===")
for i, part in enumerate(parts):
    s_type = sect_types[i] if i < len(sect_types) else "unknown"
    non_empty = [p.text for p in part if p.text.strip()]
    print(f"Section {i} (type={s_type}):")
    print(f"  Total paragraphs: {len(part)}")
    print(f"  Non-empty paragraphs: {len(non_empty)}")
    if non_empty:
        print(f"  First few: {non_empty[:3]}")
    else:
        print(f"  (All paragraphs are empty!)")
print("========================")

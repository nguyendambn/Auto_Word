import docx
import re

doc = docx.Document(r'C:\Users\The Linh\Desktop\CNTT_2022607319_NguyenVanDam_Baocao.docx')
output_lines = ["--- HEADINGS IN ORIGINAL DOCUMENT (FIRST 100 PARAGRAPHS) ---"]
count = 0
for idx, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if not text:
        continue
    # Let's show paragraphs that look like headings or chapter titles
    if (p.style.name.startswith("Heading") or 
        re.match(r'^(CHƯƠNG|MỞ ĐẦU|LỜI|DANH MỤC|MỤC LỤC)\b', text, re.IGNORECASE) or
        re.match(r'^\d+(\.\d+)*\b', text)):
        output_lines.append(f"Index {idx} | Style: {p.style.name} | Text: {text}")
        count += 1
        if count > 100:
            break

with open("scratch/headings_out.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(output_lines))
print("Done writing headings to scratch/headings_out.txt")

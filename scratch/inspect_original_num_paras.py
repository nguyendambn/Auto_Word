import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\The Linh\Desktop\CNTT_2022607319_NguyenVanDam_Baocao.docx')

print("--- paragraphs containing just numbers or page number text ---")
for idx, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if text.isdigit():
        print(f"Index {idx:4d} | Text: {text}")
    elif text in ['i', 'ii', 'iii', 'iv', 'v']:
        print(f"Index {idx:4d} | Text: {text}")

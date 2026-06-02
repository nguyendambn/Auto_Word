import sys
import docx

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\The Linh\Desktop\CNTT_2022607319_NguyenVanDam_Baocao.docx')
p = doc.paragraphs[169]
print("Paragraph 169 XML:")
print(docx.oxml.xmlchemy.serialize_for_reading(p._p))

import sys
import docx
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document(r'C:\Users\The Linh\Desktop\CNTT_2022607319_NguyenVanDam_Baocao.docx')

try:
    settings_el = doc.settings.element if hasattr(doc.settings, 'element') else doc.settings._element
    even_odd = settings_el.find(qn('w:evenAndOddHeaders'))
    print("evenAndOddHeaders in settings.xml:", even_odd is not None)
    if even_odd is not None:
        print(docx.oxml.xmlchemy.serialize_for_reading(even_odd))
except Exception as e:
    print("Error:", e)

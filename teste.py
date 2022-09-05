from docx import Document
fn='REPORT_1.docx'
document = Document(fn)

my_chapters = ['1   Certificação', '1.1 Metodologia', '1.2 Diagnóstico', '1.3 Status', '2   Resultados', '3   Conclusão']

pn=1    
import re

for p in document.paragraphs:
    r=re.match(my_chapters[4], p.text)
    if r:
        print(r.group(),pn)
    for run in p.runs:
        if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
            pn+=1
            # print('!!','='*50,pn)
from docx import Document
import re

def tableOfContents(dado, doc):
    
    fn=doc
    document = Document(fn)

    pn=1
    lista = list()

    for p in document.paragraphs:
        r=re.match(dado, p.text)
        if r:
            # print(r.group(),pn)
            lista.append(pn)
        for run in p.runs:
            if 'w:br' in run._element.xml and 'type="page"' in run._element.xml:
                pn+=1
                # print('!!','='*50,pn)
    return lista[-1]

def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        inline = paragraph.runs
        for item in inline:
            if key in item.text:
                item.text = item.text.replace(key, value)

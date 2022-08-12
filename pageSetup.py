from docx import Document
from docx.shared import Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement, ns

document = Document()

class SetupPage:

    def __init__(self, document):
        self.document = document
    
    def marginsPage(self, top, left, bottom, right):
        sections = self.document.sections
        for section in sections:
            section.top_margin = Cm(top)
            section.bottom_margin = Cm(bottom)
            section.left_margin = Cm(left)
            section.right_margin = Cm(right)
    
    def create_element(self, name):
        return OxmlElement(name)

    def create_attribute(self, element, name, value):
        element.set(ns.qn(name), value)

    def add_page_number(self, run):

        run.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        run = run.add_run()

        fldChar1 = SetupPage(document).create_element('w:fldChar')
        SetupPage(document).create_attribute(fldChar1, 'w:fldCharType', 'begin')

        instrText = SetupPage(document).create_element('w:instrText')
        SetupPage(document).create_attribute(instrText, 'xml:space', 'preserve')
        instrText.text = "PAGE"

        fldChar2 = SetupPage(document).create_element('w:fldChar')
        SetupPage(document).create_attribute(fldChar2, 'w:fldCharType', 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
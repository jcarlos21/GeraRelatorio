from docx import Document
from docx.shared import Inches, Cm

document = Document()

# https://stackoverflow.com/questions/32914595/modify-docx-page-margins-with-python-docx

class SetupPage:

    def __init__(self, document):
        self.document = document
    
    def marginsPage(self, top, bottom, left, right):
        sections = self.document.sections
        for section in sections:
            section.top_margin = Cm(0.5)
            section.bottom_margin = Cm(0.5)
            section.left_margin = Cm(1)
            section.right_margin = Cm(1)
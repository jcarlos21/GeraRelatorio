from docx import Document
from docx.shared import Inches, Cm

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
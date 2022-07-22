from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Pt

document = Document()

class StylesText:

    def __init__(self, document):
        self.document = document
        pass

    def addStyles(self, r, fonte, negrito, italico, tamanhoFonte):
        r.font.name = fonte
        r.font.size = Pt(tamanhoFonte)
        r.font.bold = negrito
        r.font.italic = italico
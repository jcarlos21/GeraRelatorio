from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from stylesTexts import StylesText
from docx.shared import Inches

document = Document()
estilos = StylesText(document)

class CriaTexto:

    def __init__(self, document):
        self.document = document
        pass
    
    def addNewLine(self, qtd):
        line = self.document.add_paragraph("\n"*qtd)
        line.paragraph_format.line_spacing = 1.50
        line.paragraph_format.space_before = Pt(0)
        line.paragraph_format.space_after = Pt(0)
    
    def textoSimples (self, texto, fonte, alinhamento, negrito, italico, tam, identacao):
        """
            O alinhamento pode ser 1, 2, 3 e 4
                0 - LEFT: Left-aligned
                1 - CENTER: Center-aligned
                2 - RIGHT: Right-aligned
                3 - JUSTIFY: Fully justified.
        """

        paragrafo = self.document.add_paragraph()
        
        if identacao:
            pf = paragrafo.paragraph_format
            pf.first_line_indent = Inches(0.5)

        CriaTexto(document).textoFormat(paragrafo, alinhamento, 1.50, 0, 0)

        r = paragrafo.add_run(texto)

        estilos.addStyles(r, fonte, negrito, italico, tam)

    def addMarcadores (self, dado, fonte, alinhamento, negrito, italico, tam):
        marcador = self.document.add_paragraph()

        paragraph_format = marcador.paragraph_format
        paragraph_format.left_indent = Inches(0.5)

        CriaTexto(document).textoFormat(marcador, alinhamento, 1.50, 0, 0)

        marcador.style = 'List Bullet'

        r = marcador.add_run(dado)
        r.font.name = fonte
        r.font.size = Pt(tam)
        r.font.bold = negrito
        r.font.italic = italico
        # Pode ser fatorado
    
    def textoFormat(self, instancia, alinhamento, space, space_after, space_before):
        instancia.alignment = alinhamento
        instancia.paragraph_format.line_spacing = space
        instancia.paragraph_format.space_before = Pt(space_after)
        instancia.paragraph_format.space_after = Pt(space_before)

    def alimentaTabela (self, rowTable, listRows, fonte, tam):
        for i in range(0, len(listRows)):
            rowTable[i].text = listRows[i]
            p = rowTable[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.runs[0]
            estilos.addStyles(r, fonte, False, False, tam)

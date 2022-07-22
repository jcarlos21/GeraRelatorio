from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Pt

document = Document()

class CriaTexto:

    def __init__(self, document):
        self.document = document
        pass
    
    def addNovaSection(self):
        newSection = self.document.add_section(WD_SECTION.NEW_PAGE)
        newSection.different_first_page_header_footer = True
    
    def textoSimples (self, texto, alinhamento, negrito):
        """
            O alinhamento pode ser 1, 2, 3 e 4
                0 - LEFT: Left-aligned
                1 - CENTER: Center-aligned
                2 - RIGHT: Right-aligned
                3 - JUSTIFY: Fully justified.

            Example:

            from docx.enum.text import WD_ALIGN_PARAGRAPH

            paragraph = document.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        """

        paragrafo = self.document.add_paragraph()
        paragrafo.alignment = alinhamento
        paragrafo.paragraph_format.line_spacing = 1.50
        r = paragrafo.add_run(texto)
        r.font.name = 'Arial'
        r.font.size = Pt(12)
        if negrito: r.font.bold = True

    def criaCabecalho (self, textoCabecalho, alinhamento):
        
        section = self.document.sections[0]
        header = section.header
        # header = section.first_page_header

        cabecalho = header.paragraphs[0]
        cabecalho.text = textoCabecalho  # Para textos com quebra de linha usar """ """.
        cabecalho.alignment = alinhamento  # Siga a docstring do método textoSimples()
        cabecalho_styles = self.document.styles["Header"]
        cabecalho_styles.font.name = 'Arial'
        cabecalho_styles.font.size = Pt(12)
        
    
    def criaTitulo (self, textoTitulo, anexo):

        titulo = self.document.add_paragraph()
        titulo.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        titulo.paragraph_format.line_spacing = 1.50
        r = titulo.add_run(f"""{textoTitulo} {anexo}""")
        r.font.name = 'Arial'
        r.font.size = Pt(12)
        r.bold = True
    
    def criaRodape (self, textoRodape, data):

        section = self.document.sections[0]
        footer = section.footer
        # footer = section.first_page_footer

        footer_p = footer.paragraphs[0]
        footer_p.text = f"{textoRodape}\n{data}"
        footer_p.alignment = 1
        footer_p_styles = self.document.styles["Footer"]
        footer_p_styles.font.name = 'Arial'
        footer_p_styles.font.size = Pt(12)

        


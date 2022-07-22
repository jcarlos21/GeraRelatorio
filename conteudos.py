from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from stylesTexts import StylesText

document = Document()
estilos = StylesText(document)

class CriaTexto:

    def __init__(self, document):
        self.document = document
        pass
    
    def addNovaSection(self):
        newSection = self.document.add_section(WD_SECTION.NEW_PAGE)
        newSection.different_first_page_header_footer = True
    
    def addNewLine(self):
        line = self.document.add_paragraph("")
        line.paragraph_format.line_spacing = 1.50
        line.paragraph_format.space_before = Pt(0)
        line.paragraph_format.space_after = Pt(0)
    
    def textoSimples (self, texto, alinhamento, negrito, italico, tam):
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
        paragrafo.paragraph_format.space_before = Pt(0)
        paragrafo.paragraph_format.space_after = Pt(0)
        
        r = paragrafo.add_run(texto)

        estilos.addStyles(r, 'Arial', negrito, italico, tam)

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
        texto = f"""{textoTitulo} {anexo}"""
        r = titulo.add_run(texto)
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

    def addMarcadores (self, dado, alinhamento):
        marcador = self.document.add_paragraph()
        marcador.alignment = alinhamento
        marcador.paragraph_format.line_spacing = 1.50
        marcador.paragraph_format.space_before = Pt(0)
        marcador.paragraph_format.space_after = Pt(0)
        marcador.style = 'List Bullet'

        r = marcador.add_run(dado)
        r.font.name = 'Arial'
        r.font.size = Pt(12)
        r.font.bold = True




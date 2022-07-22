from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from conteudos import CriaTexto

document = Document()
texto = CriaTexto(document)

# ============================================= Header ===================================================== #

cabecalho = """PONTO DE PRESENÇA DA REDE NACIONAL DE ENSINO E PESQUISA NO RIO GRANDE DO NORTE - POP-RN
REDE GIGAMETROPOLE
DEPARTAMENTO DE ENGENHARIA E OPERAÇÕES"""
texto.criaCabecalho(cabecalho, 1)

# ============================================= Footer ===================================================== #

data = '05 de Julho de 2022'
rodape = "Natal/RN"
texto.criaRodape(rodape, data)

# ============================================== Autor e Título ============================================ #

autor = "\n"*5 + "José Carlos dos Santos" + "\n"*8
texto.textoSimples(autor, 1, 0)

bilhete = "2022.2-BR01"
titulo = "Rede Giga Metrópole\nRelatório de Conformidade Referente ao Bilhete"
texto.criaTitulo(titulo, bilhete)

# ===================================== Sumario =============================================== #

texto.addNovaSection()

# document.add_page_break()  # Inserting a new page
# section1 = document.sections[0]
# header1 = section1.header
# header1.is_linked_to_previous = True

autor = "\n"*10 + "Aqui será o sumário!!!" + "\n"*12  # Deve ser coletado por meio do arquivo de configuração

textoAutor = document.add_paragraph(autor)
textoAutor.alignment = 1


# ===================================== Manutenção Corretiva RGM ============================== #

texto.addNovaSection()

texto.textoSimples("Manutenção Corretiva RGM", 1, 1)




document.save(f"REPORT_{bilhete}.docx")
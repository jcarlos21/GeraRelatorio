from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from conteudos import CriaTexto

document = Document()
texto = CriaTexto(document)

# ============================================= Dados Coletados =================================================== #

# ============================================= Header&Footer ===================================================== #

cabecalho = """Ponto de Presença da Rede Nacional de Ensino e Pesquisa no Rio Grande do Norte - Pop-Rn
Rede Gigametropole
Setor de Infraestrutura"""
texto.criaCabecalho(cabecalho, 1)

data = '05 de Julho de 2022'
rodape = "Natal/RN"
texto.criaRodape(rodape, data)

# ============================================= Header&Footer ===================================================== #

# ============================================= Dado Padrão ======================================================= #


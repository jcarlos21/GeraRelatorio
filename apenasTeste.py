from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from conteudos import CriaTexto

document = Document()
texto = CriaTexto(document)

# ============================================= Dados Coletados =================================================== #

# https://pythonacademy.com.br/blog/dicts-ou-dicionarios-no-python

celularEEntidades = dict()  # irá compor as celulas e entidades
listaCelulas = dict()  # esperado que seja uma lista de dicts
listaEntidades = dict()  # esperado que cada lista de entidades seja o valor de uma celula (que é um dict)
atributosEntidades = list()  # esperado que seja os elementos abaixo:

"""
Atributos das entidades:
    - Endereço
	- Trecho (preparado pelo próprio software)
	- Range considerado
	- Potência antes da corretiva
	- Potência depois da corretiva
	- Média de potência atual
	- Imagem referente ao range (previamente salva pelo usuário na pasta do software)
"""



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


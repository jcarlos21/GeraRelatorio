from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from conteudos import CriaTexto

document = Document()
texto = CriaTexto(document)

# ============================================= Header ===================================================== #

cabecalho = """Ponto de Presença da Rede Nacional de Ensino e Pesquisa no Rio Grande do Norte - Pop-Rn
Rede Gigametropole
Setor de Infraestrutura"""
texto.criaCabecalho(cabecalho, 1)

# ============================================= Footer ===================================================== #

data = '05 de Julho de 2022'
rodape = "Natal/RN"
texto.criaRodape(rodape, data)

# ============================================== Autor e Título ============================================ #

autor = "\n"*5 + "José Carlos dos Santos" + "\n"*8
texto.textoSimples(autor, 1, False, False, 12)

bilhete = "2022.2-BR01"  # deve ser coletado no arquivo .txt
titulo = "Rede Giga Metrópole\nRelatório de Conformidade Referente ao Bilhete"
texto.criaTitulo(titulo, bilhete)

# ===================================== Sumario =============================================== #

# texto.addNovaSection()

# document.add_page_break()  # Inserting a new page
# section1 = document.sections[0]
# header1 = section1.header
# header1.is_linked_to_previous = True

# autor = "\n"*10 + "Aqui será o sumário!!!" + "\n"*12  # Deve ser coletado por meio do arquivo de configuração

# textoAutor = document.add_paragraph(autor)
# textoAutor.alignment = 1


# ===================================== Manutenção Corretiva RGM ============================== #

texto.addNovaSection()

listaCelulas = "CA2-ZN-12.1"  # deve ser coletado no arquivo .txt/ou adicionado a um dict.
objetivo = f"""Objetivo: certificar o serviço de manutenção corretiva realizado pela empresa Interjato Soluções (bilhete {bilhete}) para restabelecer à conectividade GPON na(s) célula(s) {listaCelulas}. Os dados apresentados nesse documento foram obtidos a partir do monitoramento da rede GPON realizado pelo software GRAFANA. """
textoEntidades = "Entidade(s) afetada(s) pelo rompimento do cabo de fibras óptica:"
entidades = "EM CMEI EVANGELINA ELITA DE SOUZA"  # deve ser coletado no arquivo .txt
texto.textoSimples("Manutenção Corretiva RGM", 1, True, False, 12)
texto.addNewLine()
texto.textoSimples(objetivo, 3, False, False, 12)
texto.addNewLine()
texto.textoSimples(textoEntidades, 3, False, False, 12)
texto.addNewLine()
texto.addMarcadores(entidades, 0)  # Dentro do método, use um for para mais escolas
texto.addNewLine()
textoLocalOcorrencia = "Local da Ocorrência:"  # Use um for para imprimir os locais, incluíndo os marcadores.
texto.textoSimples(textoLocalOcorrencia, 3, False, False, 12)
texto.addNewLine()
texto
# texto.textoSimples('Grupo ', 3, False, False, 12)  # Use um for para imprimir os locais, incluíndo os marcadores.
# texto.textoSimples(f'{listaCelulas}:', 3, True, False, 12)
texto.addNewLine()



# https://python-docx.readthedocs.io/en/latest/user/quickstart.html


document.save(f"REPORT_{bilhete}.docx")
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from conteudos import CriaTexto
from stylesTexts import StylesText

document = Document()
texto = CriaTexto(document)
estilos = StylesText(document)

# Páginas

# ================================================ Capa ====================================================== #

nomeTecnico = '*Nome do Técnico*'
matriculaTecnico = '****'
nomeBolsista = '*Nome do bolsista*'
matriculaBolsista = '*Matrícula do bolsista*'
bilhete = '20XX.X-BRXX'
data = 'XX de mês de 20XX'

texto.textoSimples(nomeTecnico, 1, False, False, 12)
texto.textoSimples(f'Matrícula: {matriculaTecnico}', 1, False, False, 12)
texto.textoSimples(f'Bolsista: {nomeBolsista}', 1, False, False, 12)
texto.textoSimples(f'Matrícula: {matriculaBolsista}', 1, False, False, 12)
texto.addNewLine(5)

titulo = f'Rede Giga Metrópole\nRelatório de Conformidade Referente ao Bilhete {bilhete}'
texto.textoSimples(titulo, 1, True, False, 12)
texto.addNewLine(5)

cabecalho = """Ponto de Presença da Rede Nacional de Ensino e Pesquisa no Rio Grande do Norte - POP-RN
Rede GigaMetropole
Setor de Infraestrutura"""
texto.textoSimples(cabecalho, 1, False, False, 12)
texto.addNewLine(5)

texto.textoSimples('Natal - RN', 1, False, False, 12)
texto.textoSimples(data, 1, False, False, 12)


# ============================================== Sumário ===================================================== #

# ===================================== Manutenção Corretiva RGM ============================================= #
document.add_page_break()

celulas = '*Nome da caixa*'

p = document.add_paragraph('Objetivo: certificar o serviço de manutenção corretiva realizado pela empresa Interjato Soluções (')
b = p.add_run(f'bilhete {bilhete}')
b.font.name = 'Arial'
b.font.size = Pt(12)
b.bold = True
p.add_run(') para restabelecer à conectividade GPON na(s) célula(s) ').font.name = 'Arial'
c = p.add_run(f'{celulas}. ')
c.font.name = 'Arial'
c.font.size = Pt(12)
c.bold = True
p.add_run('Os dados apresentados nesse documento foram obtidos a partir do monitoramento da rede GPON realizado pelo software ').font.name = 'Arial'
r = p.add_run('Grafana')
# r.bold = True
# r.italic = True

estilos.addStyles(r, 'Arial', True, True, 12)  # aplique esta função nas outras linhas.

p.alignment = 3
p.paragraph_format.line_spacing = 1.50
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)


# ===================================== Armazenamento do arquivo ============================================= #
document.save(f"REPORT_{bilhete}.docx")
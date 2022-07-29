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

texto.textoSimples(nomeTecnico, 'Arial', 1, False, False, 12)
texto.textoSimples(f'Matrícula: {matriculaTecnico}', 'Arial', 1, False, False, 12)
texto.textoSimples(f'Bolsista: {nomeBolsista}', 'Arial', 1, False, False, 12)
texto.textoSimples(f'Matrícula: {matriculaBolsista}', 'Arial', 1, False, False, 12)
texto.addNewLine(6)

titulo = f'Rede Giga Metrópole\nRelatório de Conformidade Referente ao Bilhete {bilhete}'
texto.textoSimples(titulo, 'Arial', 1, True, False, 12)
texto.addNewLine(6)

cabecalho = """Ponto de Presença da Rede Nacional de Ensino e Pesquisa no Rio Grande do Norte - POP-RN
Rede GigaMetropole
Setor de Infraestrutura"""
texto.textoSimples(cabecalho, 'Arial', 1, False, False, 12)
texto.addNewLine(5)

texto.textoSimples('Natal - RN', 'Arial', 1, False, False, 12)
texto.textoSimples(data, 'Arial', 1, False, False, 12)


# ============================================== Sumário ===================================================== #

# ===================================== Manutenção Corretiva RGM ============================================= #
document.add_page_break()

texto.textoSimples ('Manutenção Corretiva RGM', 'Arial', 1, True, False, 12)
texto.addNewLine(0)

celulas = '*Nome da caixa*'

p = document.add_paragraph()

t1 = p.add_run('Objetivo: certificar o serviço de manutenção corretiva realizado pela empresa Interjato Soluções (bilhete ')
estilos.addStyles(t1, 'Arial', False, False, 12)

t2 = p.add_run(f'{bilhete}')
estilos.addStyles(t2, 'Arial', True, False, 12)

t3 = p.add_run(') para restabelecer à conectividade GPON na(s) célula(s) ')
estilos.addStyles(t3, 'Arial', False, False, 12)

t4 = p.add_run(f'{celulas}')
estilos.addStyles(t4, 'Arial', True, False, 12)

t5 = p.add_run('. Os dados apresentados nesse documento foram obtidos a partir do monitoramento da rede GPON realizado pelo software ')
estilos.addStyles(t5, 'Arial', False, False, 12)

t6 = p.add_run('Grafana.')
estilos.addStyles(t6, 'Arial', True, True, 12)
texto.addNewLine(0)

texto.textoSimples ('Entidade(s) afetada(s) pelo rompimento do cabo de fibras óptica:', 'Arial', 3, False, False, 12)
texto.addNewLine(0)

entidades = ['Escola 01', 'Escola 02', 'Escola 03']
for escola in entidades:
    texto.addMarcadores(escola, 'Arial', 0, True, False, 12)

texto.addNewLine(0)


p.alignment = 3
p.paragraph_format.line_spacing = 1.50
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)

# ===================================== Armazenamento do arquivo ============================================= #
document.save(f"REPORT_{bilhete}.docx")
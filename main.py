from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from conteudos import CriaTexto
from stylesTexts import StylesText
from docx.shared import Inches

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

celulas = '*Nome da caixa*'  # pode ser uma lista

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

entidades = ['Entidade 01', 'Entidade 02', 'Entidade 03']
for escola in entidades:
    texto.addMarcadores(escola, 'Arial', 0, True, False, 12)
texto.addNewLine(0)

texto.textoSimples('Local da Ocorrência:', 'Arial', 3, False, False, 12)
texto.addNewLine(0)

enderecosEntidade = ['Rua do Bambelô - Lagoa Azul, Natal - RN', 'Rua do Fandango, 3145 - Lagoa Azul, Natal - RN', 'Rua das Crendices, 1001 - Lagoa Azul, Natal - RN']
for i in range(0, len(enderecosEntidade)):
    texto.addMarcadores(f'Endereço {i+1}: {enderecosEntidade[i]}', 'Arial', 0, False, False, 12)

q = document.add_paragraph()  # necessário pois, ao usar o método .textoSimples(), um novo .add_paragraph() é iniciado.

pf = q.paragraph_format
pf.left_indent = Inches(0.5)
r = q.add_run('Trecho(s): ')
q.style = 'List Bullet'
r.font.name = 'Arial'
r.font.size = Pt(12)

for i in range(0, len(entidades)):
    estilos.addStyles(q.add_run(f'{celulas} - {entidades[i]}'), 'Arial', False, False, 12)
    if i + 1 < len(entidades):
        estilos.addStyles(q.add_run('; '), 'Arial', False, False, 12)
texto.addNewLine(0)

causaCorrecao =  'O rompimento nas fibras foi causado por acidente por árvores.'
texto.addNewLine(0)
texto.textoSimples('Informações do Cabo:', 'Arial', 3, False, False, 12)
texto.addMarcadores(causaCorrecao, 'Arial', 0, False, False, 12)

p.alignment = 3
p.paragraph_format.line_spacing = 1.50
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)

# ===================================== Armazenamento do arquivo ============================================= #
document.save(f"REPORT_{bilhete}.docx")
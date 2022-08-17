from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from conteudos import CriaTexto
from stylesTexts import StylesText
from docx.shared import Inches
from pageSetup import SetupPage
from analysisFunctions import AnalysisFunc
from dictFill import WriteDict
from docx.oxml import OxmlElement, ns

document = Document()
texto = CriaTexto(document)
estilos = StylesText(document)
preencheDict = WriteDict()
pageConfig = SetupPage(document)

# Páginas

# ================================= Dados para alimentação do relatório ====================================== #

bilhete = '20XX.X-BRXX'
celulas = '*Nome da caixa*'
entidades = ['Entidade 01', 'Entidade 02', 'Entidade 03']
enderecosEntidade = ['Rua, Numero, Bairro, Cidade/Estado', 'Rua, Numero, Bairro, Cidade/Estado', 'Rua, Numero, Bairro, Cidade/Estado']
causaCorrecao =  'O rompimento nas fibras foi causado por acidente por árvores.'
potenciaMedia = [-13.54, -13.44, -17.39]
potMediaBeforeAfter = {entidades[0]: [-13.54, -17.65, -12.65], entidades[1]: [-13.44, -17.65, -12.65], entidades[2]: [-17.39, -17.65, -12.65]}  # pode ser alimentado via for
rangeTeste = 30

celula = 'CAIXA1'
entidade = 'ENTIDADE 1 DA CAIXA 1'
endereco = 'Rua, Numero, Bairro, Cidade/Estado'
potMedia = -13
potBefore = -14
potAfter = -12
dadosDict = dict()

dadosDict = preencheDict.fillDict(dadosDict, celula, entidade, endereco, potMedia, potBefore, potAfter)  # dever ser chamada desta forma no botão da interface

celula = 'CAIXA1'
entidade = 'ENTIDADE 2 DA CAIXA 1'
endereco = 'Rua, Numero, Bairro, Cidade/Estado'
potMedia = -11
potBefore = -15
potAfter = -11

dadosDict = preencheDict.fillDict(dadosDict, celula, entidade, endereco, potMedia, potBefore, potAfter)

celula = 'CAIXA2'
entidade = 'ENTIDADE 1 DA CAIXA 2'
endereco = 'Rua, Numero, Bairro, Cidade/Estado'
potMedia = -10
potBefore = -12
potAfter = -11

dadosDict = preencheDict.fillDict(dadosDict, celula, entidade, endereco, potMedia, potBefore, potAfter)

print(dadosDict)

# ============================================== Margins ===================================================== #

pageConfig.marginsPage(3.0, 3.0, 2.0, 2.0)

# https://stackoverflow.com/questions/56658872/add-page-number-using-python-docx

# ================================================ Capa ====================================================== #

nomeTecnico = '*Nome do Técnico*'
matriculaTecnico = '****'
nomeBolsista = '*Nome do bolsista*'
matriculaBolsista = '*Matrícula do bolsista*'
data = 'XX de mês de 20XX'

texto.textoSimples(nomeTecnico, 'Arial', 1, False, False, 12, False)
texto.textoSimples(f'Matrícula: {matriculaTecnico}', 'Arial', 1, False, False, 12, False)
texto.textoSimples(f'Bolsista: {nomeBolsista}', 'Arial', 1, False, False, 12, False)
texto.textoSimples(f'Matrícula: {matriculaBolsista}', 'Arial', 1, False, False, 12, False)
texto.addNewLine(6)

titulo = f'Rede Giga Metrópole\nRelatório de Conformidade Referente ao Bilhete {bilhete}'
texto.textoSimples(titulo, 'Arial', 1, True, False, 12, False)
texto.addNewLine(6)

cabecalho = """Ponto de Presença da Rede Nacional de Ensino e Pesquisa no Rio Grande do Norte - POP-RN
Rede GigaMetropole
Setor de Infraestrutura"""
texto.textoSimples(cabecalho, 'Arial', 1, False, False, 12, False)
texto.addNewLine(5)

texto.textoSimples('Natal - RN', 'Arial', 1, False, False, 12, False)
texto.textoSimples(data, 'Arial', 1, False, False, 12, False)

# ============================================== Sumário ===================================================== #

texto.textoSimples('INSIRA O SUMÁRIO AQUI', 'Arial', 1, False, False, 40, False)

# ================================================ Page Number =============================================== #

section = document.sections[0]
footer = section.footer
texto.addNovaSection()
pageConfig.add_page_number(footer.paragraphs[0])

# ===================================== Manutenção Corretiva RGM ============================================= #
document.add_page_break()

texto.textoSimples ('Manutenção Corretiva RGM', 'Arial', 1, True, False, 12, False)
texto.addNewLine(0)

p = document.add_paragraph()

t1 = p.add_run('Objetivo: certificar o serviço de manutenção corretiva realizado pela empresa Interjato Soluções (bilhete ')
estilos.addStyles(t1, 'Arial', False, False, 12)
estilos.addStyles(p.add_run(f'{bilhete}'), 'Arial', True, False, 12)

t3 = p.add_run(') para restabelecer à conectividade GPON na(s) célula(s) ')
estilos.addStyles(t3, 'Arial', False, False, 12)
estilos.addStyles(p.add_run(f'{celulas}'), 'Arial', True, False, 12)

t5 = p.add_run('. Os dados apresentados nesse documento foram obtidos a partir do monitoramento da rede GPON realizado pelo software ')
estilos.addStyles(t5, 'Arial', False, False, 12)
estilos.addStyles(p.add_run('Grafana.'), 'Arial', True, True, 12)
texto.addNewLine(0)

texto.textoSimples ('Entidade(s) afetada(s) pelo rompimento do cabo de fibras óptica:', 'Arial', 3, False, False, 12, False)
texto.addNewLine(0)

print(list(dadosDict.keys())) #---------------------

for caixa in dadosDict.keys():
    for escola in dadosDict[caixa].keys():
        texto.addMarcadores(escola, 'Arial', 0, True, False, 12)
texto.addNewLine(0)

texto.textoSimples('Local da Ocorrência:', 'Arial', 3, False, False, 12, False)
texto.addNewLine(0)

i=0
for caixa in dadosDict.keys():  # Lembre-se de criar um método para tratar os loops
    for escola in dadosDict[caixa].keys():
        texto.addMarcadores(f'Endereço {i+1}: {dadosDict[caixa][escola][0]}', 'Arial', 0, False, False, 12)

q = document.add_paragraph()  # necessário pois, ao usar o método .textoSimples(), um novo .add_paragraph() é iniciado.
pf = q.paragraph_format
pf.left_indent = Inches(0.5)
r = q.add_run('Trecho(s): ')
q.style = 'List Bullet'
# r.font.name = 'Arial'
# r.font.size = Pt(12)
# O trecho acima pode ser fatorado
estilos.addStyles(r, 'Arial', False, False, 12)

for caixa in dadosDict.keys():
    for escola in dadosDict[caixa].keys():
        estilos.addStyles(q.add_run(f'{caixa.upper()} - {escola.upper()};'), 'Arial', True, False, 12)
        # if i + 1 < len(entidades):
        #     estilos.addStyles(q.add_run('; '), 'Arial', True, False, 12)


# for i in range(0, len(entidades)):
#     estilos.addStyles(q.add_run(f'{celulas.upper()} - {entidades[i].upper()}'), 'Arial', True, False, 12)
#     if i + 1 < len(entidades):
#         estilos.addStyles(q.add_run('; '), 'Arial', True, False, 12)
texto.addNewLine(0)

texto.textoSimples('Informações do Cabo:', 'Arial', 3, False, False, 12, False)
texto.addNewLine(0)
texto.addMarcadores(causaCorrecao, 'Arial', 0, False, False, 12)

texto.textoFormat(p, 3, 1.50, 0, 0)
texto.textoFormat(q, 3, 1.50, 0, 0)

# ============================================== Seção 1 ===================================================== #

document.add_page_break()

texto1 = """Todos os ativos GPON da Rede Gigametrópole são monitorados pelo software GRAFANA. Dentre os parâmetros monitorados, são de interesse nesse processo de certificação os valores de potência óptica recebidos que são enviados periodicamente pelas ONUs. A certificação é baseadas nos seguintes requisitos:"""

texto.textoSimples('1   Certificação', 'Arial', 3, True, False, 12, False)
texto.addNewLine(0)
texto.textoSimples('1.1 Metodologia', 'Arial', 3, True, False, 12, False)
texto.addNewLine(0)

texto.textoSimples(texto1, 'Arial', 3, False, False, 12, True)
texto.addNewLine(0)

requisitos = ['Comparação entre os valores de potência recebidos antes e depois do incidente;',
'Comparação entre os valores de potência recebido em cada cliente (ONU) afetado pelo incidente e a média de potência recebida nos outros clientes da mesma célula;',
'Análise do comportamento do sinal recebido na(s) ONU(s), buscando identificar oscilações relevantes (maiores que 1 dB entorno do valor médio) na magnitude do sinal.']

for i in range(0, len(requisitos)):
    texto.addMarcadores(requisitos[i], 'Arial', 3, False, False, 12)
texto.addNewLine(0)

texto.textoSimples('1.2 Diagnóstico', 'Arial', 3, True, False, 12, False)
texto.addNewLine(0)
texto2 = """A comparação dos resultados obtidos pelo monitoramento apresentados no(s) gráfico(s) da(s) figura(s) apresentada(s) no Resultados e nos dados da tabela 3, mostram que os níveis de potência óptica recebidos na(s) ONU(s) são coerentes. São sintetizados nas tabelas 1 e 2, as respostas aos requisitos estabelecidos e o diagnóstico da manutenção corretiva. """
texto.textoSimples(texto2, 'Arial', 3, False, False, 12, True)
texto.addNewLine(0)

requisitos1 = ['R1 – Os valores de potência permaneceram na mesma ordem de grandeza antes e depois do incidente?',
f'R2 – Considerando que o valor médio de potência óptica recebida nas ONUs das escolas, nas células {celulas} é de ({potenciaMedia}) dBm, respectivamente, a potência obtida na(s) ONU(s) após o reparo estão na mesma ordem de grandeza do valor médio?',
'R3 – A oscilação no sinal recebido é aceitável?']

for i in range(0, len(requisitos)):
    texto.addMarcadores(requisitos1[i], 'Arial', 3, False, False, 12)

# ============================================== Seção 2 ===================================================== #

texto.textoSimples('Legendas das respostas aos requisitos:', 'Arial', 3, False, False, 12, False)
texto.textoSimples('1.  OK – Em conformidade;', 'Arial', 3, False, False, 12, True)
texto.textoSimples('2.  X – Não atende ao requisito.', 'Arial', 3, False, False, 12, True)
texto.addNewLine(0)

# Tabela 1: __________________________________________________________________________

texto.textoSimples('Tabela 1 – Resultado do diagnóstico', 'Arial', 1, False, False, 12, False)

table = document.add_table(rows=1, cols=4)
row = table.rows[0].cells
headerRow = ['ESCOLA', 'R1', 'R2', 'R3']
texto.alimentaTabela(row, headerRow, 'Arial', 12)

# Você pode criar uma lista para armazenas as entidade. A tabela abaixo deve ser dinâmica, ou seja, deverá avaliar os três requisitos.
dataDiagnostic = []  # Table data in a form of list
for i in range(0,3):  # o range do for pode variar em função da quantidade de entidades
    dataDiagnostic.append([f'Entidade {i+1}', 'OK', 'OK', 'OK'])

for escola, r1, r2, r3 in dataDiagnostic:  # Adding a row and then adding data in it.
    row = table.add_row().cells
    listaLinhas = [escola, r1, r2, r3]
    texto.alimentaTabela(row, listaLinhas, 'Arial', 12)

for cell1 in table.columns[0].cells:
    cell1.width = Inches(2.8)

for i in range(1, 4):
    for cell in table.columns[i].cells:
        cell.width = Inches(0.5)

table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
texto.addNewLine(0)

# Tabela 2: __________________________________________________________________________

texto.textoSimples('1.3 Status', 'Arial', 3, True, False, 12, False)
texto.addNewLine(0)
texto.textoSimples('O serviço de manutenção corretiva é qualificado conforme os status apresentados na tabela 2.', 'Arial', 3, False, False, 12, True)
texto.addNewLine(0)

texto.textoSimples('Tabela 2 – Status do serviço de manutenção corretiva', 'Arial', 1, False, False, 12, False)

table2 = document.add_table(rows=1, cols=2)
row2 = table2.rows[0].cells
headerRow2 = ['PONTO ATENDIDO', 'STATUS']
texto.alimentaTabela(row2, headerRow2, 'Arial', 12)

# A tabela abaixo deve ser dinâmica, ou seja, deverá avaliar os três requisitos.
dataStatus = []  # Table data in a form of list
for i in range(0,3):  # o range do for pode variar em função da quantidade de entidades
    dataStatus.append([f'Entidade {i+1}', 'APROVADO'])

for escola, status in dataStatus:  # Adding a row and then adding data in it.
    row2 = table2.add_row().cells
    listaLinhas2 = [escola, status]
    texto.alimentaTabela(row2, listaLinhas2, 'Arial', 12)

for cell2 in table2.columns[0].cells:
    cell2.width = Inches(2.8)

for i in range(1, 2):
    for cell in table2.columns[i].cells:
        cell.width = Inches(1.5)

table2.style = 'Table Grid'
table2.alignment = WD_TABLE_ALIGNMENT.CENTER
texto.addNewLine(0)

texto.textoSimples('2   Resultados', 'Arial', 3, True, False, 12, False)
texto.addNewLine(0)

# Tabela 3: __________________________________________________________________________

texto.textoSimples('Tabela 3 – Valor médio da potência óptica recebida nas ONUs', 'Arial', 1, False, False, 12, False)

table3 = document.add_table(rows=1, cols=3)
row3 = table3.rows[0].cells
headerRow3 = ['ESCOLA', 'PRxA[dBm]', 'PRxB[dBm]']
texto.alimentaTabela(row3, headerRow3, 'Arial', 12)

dataStatus = []  # Table data in a form of list
for i in range(0,3):  # o range do for pode variar em função da quantidade de entidades
    dataStatus.append([f'Entidade {i+1}', str(-17.65), str(-12.65)])

for escola, ptA, ptD in dataStatus:  # Adding a row and then adding data in it.
    row3 = table3.add_row().cells
    listaLinhas3 = [escola, ptA, ptD]
    texto.alimentaTabela(row3, listaLinhas3, 'Arial', 12)

for cell3 in table3.columns[0].cells:
    cell3.width = Inches(2.8)

for i in range(1, 3):
    for cell in table3.columns[i].cells:
        cell.width = Inches(1.0)

table3.style = 'Table Grid'
table3.alignment = WD_TABLE_ALIGNMENT.CENTER
texto.addNewLine(0)

texto.addMarcadores ('PRxA [dBm] - Potência óptica recebida na ONU antes do incidente;', 'Arial', 3, False, False, 12)
texto.addMarcadores ('PRxB [dBm] - Potência óptica recebida na ONU após o reparo.', 'Arial', 3, False, False, 12)

# ============================================== Seção 3 ===================================================== #
document.add_page_break()

texto3 = f'No(s) gráfico(s) apresentado(s) na(s) figura(s) a seguir, os resultados mostram o comportamento do sinal recebido durante o período de {rangeTeste} dias para as escolas {entidades}, considerando antes e após o serviço de reparação ser executado. É importante ressaltar que no decorrer do período de amostragem apresentado no(s) gráfico(s) podem ocorrer intervalos sem amostras, como o período de observação é grande e os dados são enviados pelas ONU, é possível que em algum momento o equipamento seja desligado.'
texto.textoSimples(texto3, 'Arial', 3, False, False, 12, True)
texto.addNewLine(0)

for i in range(0, len(entidades)):
    texto.textoSimples(f'Figura {i+1} - Monitoramento GPON: potência óptica recebida na ONU da {entidades[i]}', 'Arial', 1, False, False, 12, False)
    document.add_picture(f'img/{str(i+1)}.JPG')
    texto.addNewLine(0)

# ============================================== Seção 4 ===================================================== #

document.add_page_break()

texto.textoSimples('3   Conclusão', 'Arial', 3, True, False, 12, False)
texto.addNewLine(0)
texto4 = 'Conclui-se que os resultados apresentados nesse documento certificam que o serviço de manutenção corretiva foi executado em conformidade com os padrões exigidos e sanando todas as pendências, garantindo o correto funcionamento da rede.'
texto.textoSimples(texto4, 'Arial', 3, False, False, 12, True)

# ===================================== Armazenamento do arquivo ============================================= #

document.save(f"REPORT_{bilhete}.docx")  # deve ser chamada pelo botão da interface


# ========================================== Interface Gráfica =============================================== #

# from tkinter import *
# from tkinter import messagebox

# class ScreenMain:
#     def __init__(self, root):
#         self.root = root
#         self.root.title("Gerador de Relatórios - POP-RN/RNP")
#         self.root.configure(background="blue")
#         self.root.geometry("700x500")
#         self.root.iconbitmap("imgMainScreen/doc2.ico")
#         self.root.resizable(False, False)

#         # ============= Imagem de Fundo ================================== #
#         self.backGroundImage = PhotoImage(file="imgMainScreen/background1.png")
#         Label(root, image=self.backGroundImage).place(x=0, y=0)

#         # ==================== Barra de Menu ============================= #
#         self.menubar = Menu(root)
#         self.file = Menu(root, tearoff=False)
#         self.file.add_separator()
#         self.file.add_command(label="Exit", command=self.exitLogin)
#         self.menubar.add_cascade(label="File", menu=self.file)

#         self.file2 = Menu(root, tearoff=False)
#         self.file2.add_command(label="Version 1.0")
#         self.menubar.add_cascade(label="About", menu=self.file2)

#         self.root.config(menu=self.menubar)

#         # ==================== Botões ============================= #
    
#     def exitLogin(self):
#         self.result = messagebox.askquestion('System', 'Are you sure you want to exit?', icon="warning")
#         if self.result == 'yes':
#             self.root.destroy()
#             exit()


# root = Tk()
# obj = ScreenMain(root)

# if __name__ == "__main__":
#     root.mainloop()

# git config --list
# git pull


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
import datetime

document = Document()
texto = CriaTexto(document)
estilos = StylesText(document)
preencheDict = WriteDict()
pageConfig = SetupPage(document)
analise = AnalysisFunc(document)

# Páginas

# ================================= Dados para alimentação do relatório ====================================== #

bilhete = '20XX.X-BRXX'
causaCorrecao =  'O rompimento nas fibras foi causado por acidente por árvores.'
rangeTeste = 30
observations = ""

celula = 'CELULA 1'
entidade = 'ENTIDADE 1 DA CELULA 1'
endereco = 'Rua, Numero, Bairro, Cidade/Estado'
potMedia = -16
potBefore = -16
potAfter = -17
dadosDict = dict()

dadosDict = preencheDict.fillDict(dadosDict, celula, entidade, endereco, potMedia, potBefore, potAfter)  # dever ser chamada desta forma no botão da interface

celula = 'CELULA 1'
entidade = 'ENTIDADE 2 DA CELULA 1'
endereco = 'Rua, Numero, Bairro, Cidade/Estado'
potMedia = -11
potBefore = -15.80
potAfter = -12.79

dadosDict = preencheDict.fillDict(dadosDict, celula, entidade, endereco, potMedia, potBefore, potAfter)

celula = 'CELULA 2'
entidade = 'ENTIDADE 1 DA CELULA 2'
endereco = 'Rua, Numero, Bairro, Cidade/Estado'
potMedia = -10
potBefore = -12.85
potAfter = -15.64

dadosDict = preencheDict.fillDict(dadosDict, celula, entidade, endereco, potMedia, potBefore, potAfter)

print('Linha 62:', dadosDict)

# ============================================== Margins ===================================================== #

pageConfig.marginsPage(3.0, 3.0, 2.0, 2.0)

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

t1 = p.add_run('Objetivo: certificar o serviço de manutenção corretiva realizado pela empresa Interjato Soluções (')
estilos.addStyles(t1, 'Arial', False, False, 12)
estilos.addStyles(p.add_run(f'bilhete {bilhete}'), 'Arial', True, False, 12)

t3 = p.add_run(') para restabelecer à conectividade GPON na(s) célula(s) ')
estilos.addStyles(t3, 'Arial', False, False, 12)
texto.repeteListaEmUmaLinha(list(dadosDict.keys()), p, 'Arial', True, False, 12)

t5 = p.add_run('. Os dados apresentados nesse documento foram obtidos a partir do monitoramento da rede GPON realizado pelo software ')
estilos.addStyles(t5, 'Arial', False, False, 12)
estilos.addStyles(p.add_run('Grafana.'), 'Arial', True, True, 12)
texto.addNewLine(0)

texto.textoSimples ('Entidade(s) afetada(s) pelo rompimento do cabo de fibras óptica:', 'Arial', 3, False, False, 12, False)
texto.addNewLine(0)

texto.repetMarcadoreEntidade(dadosDict, 'Arial', 0, negrito=True, italico=False, tam=12)
texto.addNewLine(0)

texto.textoSimples('Local da Ocorrência:', 'Arial', 3, False, False, 12, False)
texto.addNewLine(0)

texto.imprimeMarcadorEndereco(dadosDict, 'Arial', negrito=False, italico=False, tam=12)

q = document.add_paragraph()  # necessário pois, ao usar o método .textoSimples(), um novo .add_paragraph() é iniciado.
pf = q.paragraph_format
pf.left_indent = Inches(0.5)
q.style = 'List Bullet'
texto.repeteListaEmUmaLinha('Trecho(s): ', q, fonte='Arial', negrito=False, italico=False, tam=12)

texto.imprimeCaixaEntidade(dadosDict, q, fonte='Arial', negrito=True, italico=False, tam=12)
texto.addNewLine(0)

texto.textoSimples('Informações do Cabo:', 'Arial', 3, False, False, 12, False)
texto.addNewLine(0)
texto.addMarcadores(causaCorrecao, 'Arial', 0, False, False, 12)

texto.textoFormat(p, 3, 1.50, 0, 0)
texto.textoFormat(q, 0, 1.50, 0, 0)

# ============================================== Seção 1 ===================================================== #

document.add_page_break()

texto1 = """Todos os ativos GPON da Rede Gigametrópole são monitorados pelo software GRAFANA. Dentre os parâmetros monitorados, são de interesse nesse processo de certificação os valores de potência óptica recebidos que são enviados periodicamente pelas ONU(s). A certificação é baseadas nos seguintes requisitos:"""

texto.textoSimples('1   Certificação', 'Arial', 3, True, False, 12, False)
texto.addNewLine(0)
texto.textoSimples('1.1 Metodologia', 'Arial', 3, True, False, 12, False)
texto.addNewLine(0)

texto.textoSimples(texto1, 'Arial', 3, False, False, 12, True)
texto.addNewLine(0)

requisitos = ['Comparação entre os valores de potência recebidos antes e depois do incidente;',
'Comparação entre os valores de potência recebido em cada cliente (ONU) afetado pelo incidente e a média de potência recebida nos outros clientes da mesma célula;',
'Análise do comportamento do sinal recebido na(s) ONU(s), buscando identificar oscilações relevantes (maiores que 1 dB entorno do valor médio) na magnitude do sinal.']

texto.repetMarcadores(requisitos, 'Arial', 3, False, False, 12)
texto.addNewLine(0)

texto.textoSimples('1.2 Diagnóstico', 'Arial', 3, True, False, 12, False)
texto.addNewLine(0)
texto2 = """A comparação dos resultados obtidos pelo monitoramento apresentados no(s) gráfico(s) da(s) figura(s) apresentada(s) no Resultados e nos dados da tabela 3, mostram que os níveis de potência óptica recebidos na(s) ONU(s) são coerentes. São sintetizados nas tabelas 1 e 2, as respostas aos requisitos estabelecidos e o diagnóstico da manutenção corretiva. """
texto.textoSimples(texto2, 'Arial', 3, False, False, 12, True)
texto.addNewLine(0)

# ============================================== Requisitos de teste ========================================= #

r1 = 'R1 – Os valores de potência permaneceram na mesma ordem de grandeza antes e depois do incidente?'
texto.addMarcadores(r1, 'Arial', 3, False, False, 12)

pj = document.add_paragraph()
pk = pj.paragraph_format
pk.left_indent = Inches(0.5)
r = pj.add_run('R2 – Considerando que o valor médio de potência óptica recebida nas ONU(s) das escolas, nas células ')
pj.style = 'List Bullet'
estilos.addStyles(r, 'Arial', False, False, 12)
texto.textoFormat(pj, 3, 1.50, 0, 0)

texto.repeteListaEmUmaLinha(list(dadosDict.keys()), pj, 'Arial', True, False, 12)

r2b = ' é de '
texto.repeteListaEmUmaLinha(r2b, pj, 'Arial', False, False, 12)

for caixa in dadosDict.keys():  # Lembre-se de criar um método para tratar os loops
    for escola in dadosDict[caixa].keys():
        texto.repeteListaEmUmaLinha(str(dadosDict[caixa][escola][1]) + ' dBm, ', pj, 'Arial', False, False, 12)

r2c = ' a potência obtida na(s) ONU(s) após o reparo estão na mesma ordem de grandeza do valor médio?'
texto.repeteListaEmUmaLinha(r2c, pj, 'Arial', False, False, 12)

r3 = 'R3 – A oscilação no sinal recebido é aceitável?'
texto.addMarcadores(r3, 'Arial', 3, False, False, 12)

# ============================================== Seção 2 ===================================================== #

texto.textoSimples('Legendas das respostas aos requisitos:', 'Arial', 3, False, False, 12, False)
texto.addNewLine(0)
texto.textoSimples('1.  OK – Em conformidade;', 'Arial', 3, False, False, 12, True)
texto.textoSimples('2.  X – Não atende ao requisito.', 'Arial', 3, False, False, 12, True)
texto.addNewLine(0)

# Tabela 1: __________________________________________________________________________

texto.textoSimples('Tabela 1 – Resultado do diagnóstico', 'Arial', 1, False, False, 12, False)

table = document.add_table(rows=1, cols=4)
row = table.rows[0].cells
headerRow = ['ESCOLA', 'R1', 'R2', 'R3']
texto.alimentaTabela(row, headerRow, 'Arial', 12)

dataDiagnostic = list()
dataDiagnosticStatus = list()
dataDiagnosticPot = list()
for enS in dadosDict.keys():
    for enI in dadosDict[enS].keys():
        a1 = analise.rR1(dadosDict[enS][enI][2], dadosDict[enS][enI][3])
        a2 = analise.rR2(dadosDict[enS][enI][1], dadosDict[enS][enI][3])
        a3 = analise.rR3(dadosDict[enS][enI][3])
        status = analise.rStatus(a1, a2, a3)
        dataDiagnostic.append([enI, a1, a2, a3])
        dataDiagnosticStatus.append([enI, status])
        dataDiagnosticPot.append([enI, str(dadosDict[enS][enI][2]), str(dadosDict[enS][enI][3])])

print(dataDiagnostic)
print(dataDiagnosticStatus)
print(dataDiagnosticPot)


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

for escola, status in dataDiagnosticStatus:  # Adding a row and then adding data in it.
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

for escola, ptA, ptD in dataDiagnosticPot:  # Adding a row and then adding data in it.
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

texto3 = f'No(s) gráfico(s) apresentado(s) na(s) figura(s) a seguir, os resultados mostram o comportamento do sinal recebido durante o período de {rangeTeste} dias, considerando antes e após o serviço de reparação ser executado. É importante ressaltar que no decorrer do período de amostragem apresentado no(s) gráfico(s) podem ocorrer intervalos sem amostras, como o período de observação é grande e os dados são enviados pelas ONU, é possível que em algum momento o equipamento seja desligado.'
texto.textoSimples(texto3, 'Arial', 3, False, False, 12, True)
texto.addNewLine(0)

for i in range(0, len(dataDiagnosticStatus)):
    texto.textoSimples(f'Figura {i+1} - Monitoramento GPON: potência óptica recebida na ONU da {dataDiagnosticStatus[i][0]}', 'Arial', 1, False, False, 12, False)
    document.add_picture(f'img/0{str(i+1)}.JPG')
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

from tkinter import *
# from Tkinter import *
from tkinter import messagebox
from tkinter import ttk
# from Tkinter import messagebox

class ScreenMain:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatórios - POP-RN/RNP")
        self.root.configure(background="blue")
        self.root.geometry("900x500")
        self.root.iconbitmap("imgMainScreen/doc2.ico")
        self.root.resizable(False, False)

        # ============= Imagem de Fundo ================================== #
        self.backGroundImage = PhotoImage(file="imgMainScreen/background8.png")
        Label(self.root, image=self.backGroundImage).place(x=0, y=0)

        self.whiteScreen = Frame(self.root, bg="white")
        self.whiteScreen.place(x=150, y=40, width=700, height=400)

        self.photo = PhotoImage(file="imgMainScreen/Jump.png")
        self.labelphoto = Label(self.root, width=200, height=300, image=self.photo)
        self.labelphoto.place(x=50, y=85)
        self.labelphoto.image = self.photo

        # ==================== Barra de Menu ============================= #
        self.menubar = Menu(self.root)
        self.file = Menu(self.root, tearoff=False)
        self.file.add_separator()
        self.file.add_command(label="Exit", command=self.exitLogin)
        self.menubar.add_cascade(label="File", menu=self.file)

        self.file2 = Menu(root, tearoff=False)
        self.file2.add_command(label="Version 1.0")
        self.menubar.add_cascade(label="About", menu=self.file2)

        self.root.config(menu=self.menubar)

        # ==================== Textos e Botões ============================= #
        Label(self.whiteScreen, text="Entre com as informações", font=("times new roman", 15, "bold"), bg="white", fg="#016AFB").place(x=120, y=10)
        Label(self.whiteScreen, text="Bilhete:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=60)
        self.bilhete = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=195, y=60, width=120)

        Label(self.whiteScreen, text="Célula:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=320, y=60)
        self.celula = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=380, y=60, width=110)

        Label(self.whiteScreen, text="Data:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=495, y=60)
        self.data = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=545, y=60, width=130)

        Label(self.whiteScreen, text="Entidade:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=100)
        self.entidade = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=195, y=100, width=165)

        Label(self.whiteScreen, text="P.M.:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=365, y=100)
        self.p_A = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=415, y=100, width=50)
    
        Label(self.whiteScreen, text="P.A.:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=475, y=100)
        self.p_A = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=520, y=100, width=50)

        Label(self.whiteScreen, text="P.D.:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=580, y=100)
        self.p_D = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=625, y=100, width=50)

        Label(self.whiteScreen, text="Endereço:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=140)
        self.endereco = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=195, y=140, width=240)

        Label(self.whiteScreen, text="Testagem:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=440, y=140)

        self.stateChosen = StringVar()
        self.stateChoose = ttk.Combobox(self.whiteScreen, textvariable=self.stateChosen, width=21)
        self.stateChoose['values'] = ['10 dais', '20 dias', '30 dias', '40 dias', '50 dias', '60 dias']
        self.stateChoose.grid(column=0, row=0, padx=525, pady=140)
        self.stateChoose.current(0)

        Label(self.whiteScreen, text="Motivo:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=180)
        self.motivo = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=195, y=180, width=480)

        Label(self.whiteScreen, text="Técnico:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=220)
        self.tecnico = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=195, y=220, width=280)

        Label(self.whiteScreen, text="Matrícula:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=480, y=220)
        self.matricula_tecnico = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=560, y=220, width=115)

        Label(self.whiteScreen, text="Bolsista:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=260)
        self.bolsista = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=195, y=260, width=280)

        Label(self.whiteScreen, text="Matrícula:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=480, y=260)
        self.matricula_bolsista = Entry(self.whiteScreen, font=("times new roman", 12), bg="lightgray").place(x=560, y=260, width=115)

    def exitLogin(self):
        self.result = messagebox.askquestion('System', 'Are you sure you want to exit?', icon="warning")
        if self.result == 'yes':
            self.root.destroy()
            exit()


root = Tk()
obj = ScreenMain(root)

if __name__ == "__main__":
    root.mainloop()

# git config --list
# git pull


# d = datetime.datetime.today().strftime('%d')
# m = datetime.datetime.today().strftime('%B')
# y = datetime.datetime.today().strftime('%Y')

# nomeTecnico = '*Nome do Técnico*'
# matriculaTecnico = '****'
# nomeBolsista = '*Nome do bolsista*'
# matriculaBolsista = '*Matrícula do bolsista*'
# data = f'{d} de {m} de {y}'

# bilhete = '20XX.X-BRXX'
# causaCorrecao =  'O rompimento nas fibras foi causado por acidente por árvores.'
# rangeTeste = 30
# observations = ""

# celula = 'CELULA 1'
# entidade = 'ENTIDADE 1 DA CELULA 1'
# endereco = 'Rua, Numero, Bairro, Cidade/Estado'
# potMedia = -16
# potBefore = -16
# potAfter = -17
# dadosDict = dict()

# dadosDict = preencheDict.fillDict(dadosDict, celula, entidade, endereco, potMedia, potBefore, potAfter)  # dever ser chamada desta forma no botão da interface


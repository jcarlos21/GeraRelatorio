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
from gerador import TextGenerator

document = Document()
texto = CriaTexto(document)
estilos = StylesText(document)
preencheDict = WriteDict()
pageConfig = SetupPage(document)
analise = AnalysisFunc(document)
relatorio = TextGenerator()

# Páginas

# ================================= Dados para alimentação do relatório ====================================== #

bilheteEnt = '20XX.X-BRXX'
causaCorrecao =  'O rompimento nas fibras foi causado por acidente por árvores.'
rangeTeste = 30
observations = ""
global dadosDict
dadosDict = dict()


# # ===================================== Armazenamento do arquivo ============================================= #

# document.save(f"REPORT_{bilheteEnt}.docx")  # deve ser chamada pelo botão da interface


# # ========================================== Interface Gráfica =============================================== #

from tkinter import *
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

root = Tk()
root.title("Gerador de Relatórios - POP-RN/RNP")
root.configure(background="blue")
root.geometry("900x500")
root.iconbitmap("imgMainScreen/doc2.ico")
root.resizable(False, False)

# ============= Imagem de Fundo ================================== #

backGroundImage = PhotoImage(file="imgMainScreen/background8.png")
Label(root, image=backGroundImage).place(x=0, y=0)

whiteScreen = Frame(root, bg="white")
whiteScreen.place(x=150, y=40, width=700, height=400)

photo = PhotoImage(file="imgMainScreen/logo_poprn.png")
labelphoto = Label(root, width=200, height=300, image=photo)
labelphoto.place(x=50, y=85)
labelphoto.image = photo

# ====================== Funções de coleta de dados =============== #

def exitLogin():
    result = messagebox.askquestion('System', 'Are you sure you want to exit?', icon="warning")
    if result == 'yes':
        root.destroy()
        exit()

def add_ao_arquivo():

    preencheDict = WriteDict()

    bilheteEnt = bilhete.get()
    preencheDict.dadosDict = dadosDict
    preencheDict.celula = celula.get()
    dataEnt = data.get()
    preencheDict.entidade = entidade.get()
    preencheDict.endereco = endereco.get()
    preencheDict.potMedia = p_M.get()
    preencheDict.potBefore = p_A.get()
    preencheDict.potAfter = p_D.get()
    causaCorrecao = motivo.get()
    tecnicoEnt = tecnico.get()
    matTecnico = matricula_tecnico.get()
    bolsistaEnt = bolsista.get()
    matBolsista = matricula_bolsista.get()
    observations = obervacao.get()

    mensagem_insercao['text'] = preencheDict.fillDict()

    # Excluindo dados digitados nas caixas
    
    # bilhete.delete(0, END)
    # celula.delete(0, END)
    # data.delete(0, END)
    entidade.delete(0, END)
    endereco.delete(0, END)
    p_M.delete(0, END)
    p_A.delete(0, END)
    p_D.delete(0, END)
    motivo.delete(0, END)
    tecnico.delete(0, END)
    matricula_tecnico.delete(0, END)
    bolsista.delete(0, END)
    matricula_bolsista.delete(0, END)
    obervacao.delete(0, END)

    # return dadosDict

def gerar_arquivo():
    document.save(f"REPORT_{bilheteEnt}.docx")

    # pass

# ===================== Barra de Menu ============================= #

menubar = Menu(root)
file = Menu(root, tearoff=False)
file.add_separator()
file.add_command(label="Exit", command=exitLogin)
menubar.add_cascade(label="File", menu=file)

file2 = Menu(root, tearoff=False)
file2.add_command(label="Version 1.0")
menubar.add_cascade(label="About", menu=file2)

root.config(menu=menubar)

# ==================== Rótulos e entradas ============================= #

Label(whiteScreen, text="Entre com as informações", font=("times new roman", 15, "bold", "italic"), bg="white", fg="#016AFB").place(x=250, y=10)
Label(whiteScreen, text="Bilhete:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=60)
bilhete = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
bilhete.place(x=195, y=60, width=120)

Label(whiteScreen, text="Célula:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=320, y=60)
celula = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
celula.place(x=380, y=60, width=110)

Label(whiteScreen, text="Data:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=495, y=60)
data = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
data.place(x=545, y=60, width=130)

Label(whiteScreen, text="Escola:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=100)
entidade = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
entidade.place(x=195, y=100, width=165)

Label(whiteScreen, text="P.M.:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=365, y=100)
p_M= Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
p_M.place(x=415, y=100, width=50)

Label(whiteScreen, text="P.A.:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=475, y=100)
p_A = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
p_A.place(x=520, y=100, width=50)

Label(whiteScreen, text="P.D.:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=580, y=100)
p_D = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
p_D.place(x=625, y=100, width=50)

Label(whiteScreen, text="Endereço:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=140)
endereco = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
endereco.place(x=195, y=140, width=240)

Label(whiteScreen, text="Testagem:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=440, y=140)

stateChosen = StringVar()
stateChoose = ttk.Combobox(whiteScreen, textvariable=stateChosen, width=21)
stateChoose['values'] = ['10 dais', '20 dias', '30 dias', '40 dias', '50 dias', '60 dias']
stateChoose.grid(column=0, row=0, padx=525, pady=140)
stateChoose.current(0)

Label(whiteScreen, text="Motivo:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=180)
motivo = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
motivo.place(x=195, y=180, width=480)

Label(whiteScreen, text="Técnico:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=220)
tecnico = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
tecnico.place(x=195, y=220, width=280)

Label(whiteScreen, text="Mat. Tec.:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=480, y=220)
matricula_tecnico = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
matricula_tecnico.place(x=560, y=220, width=115)

Label(whiteScreen, text="Bolsista:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=260)
bolsista = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
bolsista.place(x=195, y=260, width=280)

Label(whiteScreen, text="Mat. Bol.:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=480, y=260)
matricula_bolsista = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
matricula_bolsista.place(x=560, y=260, width=115)

Label(whiteScreen, text="Observ.:", font=("times new roman", 12, "bold"), bg="white", fg="gray").place(x=120, y=300)
obervacao = Entry(whiteScreen, font=("times new roman", 12), bg="lightgray")
obervacao.place(x=195, y=300, width=480)

Label(root, text="Desenvolvido por:\nJosé Carlos dos Santos", font=("arial", 7, "italic"), bg="#93D9F0", fg="black", justify='left').place(x=790, y=450)

# ==================== Botões ============================= #

botao_add_arquivo = Button(whiteScreen, text="Adicionar", borderwidth=3, cursor="hand2")  # Definindo botão salvar
botao_add_arquivo['command'] = add_ao_arquivo    # trata-se de uma função mais adiante
botao_add_arquivo.place(x=520, y=360)

botao_salvar_doc = Button(whiteScreen, text="Gerar arquivo", bg="blue", fg="white", borderwidth=3, cursor="hand2")
botao_salvar_doc['command'] = gerar_arquivo
botao_salvar_doc.place(x=590, y=360)

# ==================== Mensagens de confirmação ============================= #

mensagem_insercao =  Label(whiteScreen, text="", font=("arial", 10, "italic"), bg="white", fg="green")
mensagem_insercao.place(x=275, y=345)
mensagem_gravacao =  Label(whiteScreen, font=("arial", 10, "italic"), bg="white", fg="green")
mensagem_gravacao.place(x=275, y=350)

print(dadosDict)

if __name__ == "__main__":
    root.mainloop()

# git config --list
# git pull


# d = datetime.datetime.today().strftime('%d')
# m = datetime.datetime.today().strftime('%B')
# y = datetime.datetime.today().strftime('%Y')

# data = f'{d} de {m} de {y}'
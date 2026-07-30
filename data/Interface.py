import customtkinter as ctk
import logging
from data.main import main
import os

logging.basicConfig(
filename='aplicacao.log',
filemode='w',
level=logging.DEBUG,
format='%(asctime)s - %(levelname)s - %(message)s'
)

app = ctk.CTk()
app.geometry('400x300')
app.title('Automação')

def conc():
    try:
        with open("dados.txt", "r", encoding="utf-8") as f:
            conteudo = f.read()
            if conteudo == "Arquivo processado com sucesso!":
                label_concluido.configure(text=conteudo)
    except Exception as e:
        label_concluido.configure(text=f"Erro ao ler o arquivo: {e}")

instr = ctk.CTkLabel(app, text='Clique no botão para tratar o arquivo Excel')
instr.pack(pady=20)

def funcoes():
    try:
        main()
        if not os.path.exists("dados.txt"):
            with open("dados.txt", "w", encoding="utf-8") as f:
                f.write("Arquivo processado com sucesso!")
        conc()
    except Exception as e:
        label_concluido.configure(text=f"Erro ao executar o script: {e}")

botao = ctk.CTkButton(app, text='Tratar arquivo', command=funcoes)
botao.pack(pady=20)

label_concluido = ctk.CTkLabel(app, text='')
label_concluido.pack(pady=20)


app.mainloop()
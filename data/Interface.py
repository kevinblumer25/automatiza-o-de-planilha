import logging
import os
import sys
from pathlib import Path

import customtkinter as ctk

try:
    from main import main
except ImportError:
    from main import main


def get_app_dir():
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        if exe_dir and os.path.isdir(exe_dir):
            return exe_dir
        return os.getcwd()
    return os.path.dirname(os.path.abspath(__file__))


def get_work_dir():
    return Path(get_app_dir())


app_dir = get_work_dir()
app_dir.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    filename=os.path.join(app_dir, 'aplicacao.log'),
    filemode='w',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

app = ctk.CTk()
app.geometry('400x300')
app.title('Automação')

def conc():
    dados_path = os.path.join(app_dir, "dados.txt")
    try:
        with open(dados_path, "r", encoding="utf-8") as f:
            conteudo = f.read()
            if conteudo == "Arquivo processado com sucesso!":
                label_concluido_cart.configure(text=conteudo)
    except Exception as e:
        label_concluido_cart.configure(text=f"Erro ao ler o arquivo: {e}")

instr = ctk.CTkLabel(app, text='Clique no botão para tratar o arquivo Excel')
instr.pack(pady=20)

def funcoes():
    try:
        main()
        dados_path = os.path.join(app_dir, "dados.txt")
        if not os.path.exists(dados_path):
            with open(dados_path, "w", encoding="utf-8") as f:
                f.write("Arquivo processado com sucesso!")
        conc()
    except Exception as e:
        label_concluido_cart.configure(text=f"Erro ao executar o script: {e}")

botao = ctk.CTkButton(app, text='Tratar arquivo', command=funcoes)
botao.pack(pady=20)

label_concluido_cart = ctk.CTkLabel(app, text='')
label_concluido_cart.pack(pady=20)

def funccart():
    try:
        from followup import mainFup
        mainFup()
        label_concluido_fup.configure(text="Follow Ups salvos com sucesso!")
        conc()
    except Exception as e:
        label_concluido_fup.configure(text=f"Erro ao executar o script: {e}")

label_instr_fup = ctk.CTkLabel(app, text='Clique no botão para gerar os Follow Ups')
label_instr_fup.pack(pady=20)

botao_fup = ctk.CTkButton(app, text='Gerar Follow Ups', command=funccart)
botao_fup.pack(pady=20)

label_concluido_fup = ctk.CTkLabel(app, text='')
label_concluido_fup.pack(pady=20)

app.mainloop()
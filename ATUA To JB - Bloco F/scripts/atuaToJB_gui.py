# ------------------------
# importações
# ------------------------

import tkinter as tk
from tkinter import filedialog, messagebox
import os

from atuaToJb import processar_arquivo

VERSAO = "1.0"

#------------------------------------
#   Valida se o arquivo possui layout ATUA esperado.
#   Retorna True se válido, levanta Exception se inválido.
#-------------------------------------

def validar_layout_arquivo(caminho_arquivo):
    extensoes_permitidas = (".seq", ".txt")

    if not caminho_arquivo.lower().endswith(extensoes_permitidas):
        raise Exception(
            "Arquivo inválido. \n\n"
            "Selecione um arquivo com extensão (.seq) ou (.txt)."
        )

    with open(caminho_arquivo, "r", encoding="latin-1") as f:
        linhas = f.readlines()

    if not linhas:
        raise Exception("O arquivo está vazio.")

    possui_00 = False
    possui_06 = False

    for linha in linhas:
        linha = linha.strip()

        if linha.startswith("00;"):
            possui_00 = True
        elif linha.startswith("06;"):
            possui_06 = True

        if possui_00 and possui_06:
            break

    if not possui_00:
        raise Exception("Layout inválido: nenhuma linha do tipo 00 encontrada.")

    if not possui_06:
        raise Exception("Layout inválido: nenhuma linha do tipo 06 encontrada.")

    return True

def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo ATUA (.seq)",
        filetypes=[
            ("Todos os arquivos", "*.*"),
            ("Arquivos SEQ", "*.seq"), 
            ("Arquivos Texto", "*.txt")
            
            ]
    )

    if not arquivo:
        return

    try:

        validar_layout_arquivo(arquivo)

        # Interface em modo processamento
        botao.config(state="disabled")
        status_label.config(text="Processando arquivo, aguarde...")
        janela.update_idletasks()

        processar_arquivo(arquivo)

        status_label.config(text="Processamento concluído com sucesso!")

        resposta = messagebox.askyesno(
            "Conversão concluída",
            "Arquivo processado com sucesso!\n\n"
            "Deseja converter outro arquivo?"
        )

        # Se NÃO, apenas volta para a tela inicial
        if not resposta:
            status_label.config(text="Pronto para nova conversão.")

    except Exception as e:
        messagebox.showerror(
            "Erro",
            f"Ocorreu um erro ao processar o arquivo:\n\n{str(e)}"
        )
        status_label.config(text="Erro ao processar arquivo.")

    finally:
        botao.config(state="normal")


# -----------------------------
# Janela principal
# -----------------------------

janela = tk.Tk()
janela.title("ATUA → JB")
janela.geometry("460x220")
janela.resizable(False, False)

label = tk.Label(
    janela,
    text="Conversão ATUA → JB\n\n"
         "Selecione o arquivo (.seq / .txt) de origem",
    font=("Arial", 11),
    justify="center"
)
label.pack(pady=20)

botao = tk.Button(
    janela,
    text="Selecionar arquivo",
    font=("Arial", 11),
    width=30,
    height=2,
    command=selecionar_arquivo
)
botao.pack(pady=5)

status_label = tk.Label(
    janela,
    text="Pronto para nova conversão.",
    font=("Arial", 9),
    fg="blue"
)
status_label.pack(pady=8)

rodape = tk.Label(
    janela,
    text=f"Versão {VERSAO}",
    font=("Arial", 9),
    fg="gray"
)
rodape.pack(side="bottom", pady=6)

janela.mainloop()

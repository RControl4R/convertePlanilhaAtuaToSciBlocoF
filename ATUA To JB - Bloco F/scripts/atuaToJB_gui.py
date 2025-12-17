# ------------------------
# importações
# ------------------------

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

# Importa a função principal do script atuaToJb.py
from atuaToJb import processar_arquivo


def selecionar_arquivo():
    while True:
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo ATUA (.seq)",
            filetypes=[("Arquivos SEQ", "*.seq"), ("Arquivos Texto", "*.txt")]
        )

        if not arquivo:
            # Usuário cancelou a seleção
            return

        try:
            processar_arquivo(arquivo)

            resposta = messagebox.askyesno(
                "Conversão concluída",
                "Arquivo processado com sucesso!\n\n"
                "Deseja converter outro arquivo?"
            )

            if not resposta:
                return

            # Se respondeu SIM, o loop continua

        except Exception as e:
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao processar o arquivo:\n\n{str(e)}"
            )
            return


# -----------------------------
# Janela principal
# -----------------------------
janela = tk.Tk()
janela.title("ATUA → JB")
janela.geometry("460x200")
janela.resizable(False, False)

label = tk.Label(
    janela,
    text="Conversão ATUA → JB\n\n"
         "Selecione o arquivo (.seq) de origem",
    font=("Arial", 11),
    justify="center"
)
label.pack(pady=25)

botao = tk.Button(
    janela,
    text="Selecionar arquivo",
    font=("Arial", 11),
    width=30,
    height=2,
    command=selecionar_arquivo
)
botao.pack(pady=10)

rodape = tk.Label(
    janela,
    text="Versão 1.0",
    font=("Arial", 9),
    fg="gray"
)
rodape.pack(side="bottom", pady=8)

janela.mainloop()

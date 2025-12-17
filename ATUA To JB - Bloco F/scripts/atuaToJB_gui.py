# ------------------------
# importações
# ------------------------

import tkinter as tk
from tkinter import filedialog, messagebox
import os

from atuaToJb import processar_arquivo

VERSAO = "1.0"


def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo ATUA (.seq)",
        filetypes=[("Arquivos SEQ", "*.seq"), ("Arquivos Texto", "*.txt")]
    )

    if not arquivo:
        return

    try:
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
         "Selecione o arquivo (.seq) de origem",
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

#------------------------
# importações
#------------------------

import tkinter as tk
from tkinter import filedialog, messagebox
from atuaToTxtV2 import processar_arquivo

def selecionar_arquivo():
    while True:
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xls *.xlsx")]
        )

        if not arquivo:
            return

        try:
            processar_arquivo(arquivo)

            resposta = messagebox.askyesno(
                "Conversão concluída",
                "Arquivo processado com sucesso!\n\nDeseja converter outro arquivo?"
            )

            if not resposta:
                janela.destroy()
                return

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
janela.title("Conversor ATUA → SCI")
janela.geometry("460x200")
janela.resizable(False, False)

label = tk.Label(
    janela,
    text="Conversão ATUA → SCI (Bloco F)\n\n"
         "Selecione o arquivo (.xls / .xlsx) de origem",
    font=("Arial", 11),
    justify="center"
)
label.pack(pady=25)

btn = tk.Button(
    janela,
    text="Selecionar arquivo Excel",
    command=selecionar_arquivo,
    font=("Arial", 11),
    width=30,
    height=2
)
btn.pack(expand=True)

rodape = tk.Label(
    janela,
    text="Versão 1.04",
    font=("Arial", 9),
    fg="gray"
)
rodape.pack(side="bottom", pady=8)

janela.mainloop()

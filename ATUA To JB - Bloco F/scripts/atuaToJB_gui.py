import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox


def processar_arquivo(caminho_arquivo):
    """
    Processa o arquivo .seq:
    - Remove linhas 00 com 'Pedágio'
    - Ajusta campos da linha 06
    - Gera arquivo _JB na mesma pasta
    """

    pasta = os.path.dirname(caminho_arquivo)
    nome_base = os.path.splitext(os.path.basename(caminho_arquivo))[0]
    arquivo_saida = os.path.join(pasta, f"{nome_base}_JB.seq")

    linhas_saida = []

    # Variáveis de contexto do último bloco 00 válido
    campo_02_00 = None
    campo_03_00 = None
    campo_16_00 = None

    with open(caminho_arquivo, "r", encoding="latin-1") as f:
        linhas = f.readlines()

    for linha in linhas:
        linha = linha.rstrip("\n")
        colunas = linha.split(";")

        tipo = colunas[0]

        # ======================
        # LINHA 00
        # ======================
        if tipo == "00":
            campo_13 = colunas[12] if len(colunas) > 12 else ""

            # Regra: excluir Pedágio
            if "pedágio" in campo_13.lower():
                continue

            campo_02_00 = colunas[1]
            campo_03_00 = colunas[2]
            campo_16_00 = colunas[15] if len(colunas) > 15 else ""

            linhas_saida.append(linha)
            continue

        # ======================
        # LINHA 06
        # ======================
        if tipo == "06":
            nova_linha = []

            for i, valor in enumerate(colunas):
                # Excluir coluna 23 (índice 22)
                if i == 22:
                    continue
                nova_linha.append(valor)

            # Inserções conforme regra
            # 06;2;<campo02_00>;<campo03_00>;...
            nova_linha.insert(2, campo_02_00 or "")
            nova_linha.insert(3, campo_03_00 or "")

            # Campo 19 vazio => ;"";
            nova_linha.insert(18, '""')

            # Campo 23 recebe campo 16 da linha 00
            if len(nova_linha) > 22:
                nova_linha[22] = campo_16_00 or ""

            linhas_saida.append(";".join(nova_linha))
            continue

        # ======================
        # OUTRAS LINHAS
        # ======================
        linhas_saida.append(linha)

    with open(arquivo_saida, "w", encoding="latin-1") as f:
        for l in linhas_saida:
            f.write(l + "\n")

    return arquivo_saida


def executar_gui():
    root = tk.Tk()
    root.withdraw()

    while True:
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo .seq",
            filetypes=[("Arquivos SEQ", "*.seq"), ("Todos os arquivos", "*.*")]
        )

        if not arquivo:
            break

        try:
            saida = processar_arquivo(arquivo)
            resposta = messagebox.askyesno(
                "Sucesso",
                f"Arquivo gerado com sucesso:\n\n{saida}\n\nDeseja converter outro arquivo?"
            )
            if not resposta:
                break

        except Exception as e:
            messagebox.showerror("Erro", str(e))
            break


if __name__ == "__main__":
    executar_gui()

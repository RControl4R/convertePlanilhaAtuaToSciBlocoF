import os
import sys
import unicodedata

import re

def contem_pedagio(texto):
    """
    Detecta a palavra PEDAGIO mesmo com:
    - acentos quebrados
    - caracteres invisíveis
    - lixo no meio da palavra
    """
    texto = texto.lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if c.isalpha())
    return "pedagio" in texto



def normalizar_texto(texto):
    texto = texto.lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto


def processar_arquivo(caminho_entrada):
    pasta, nome_arquivo = os.path.split(caminho_entrada)
    nome_base, extensao = os.path.splitext(nome_arquivo)
    caminho_saida = os.path.join(pasta, f"{nome_base}_JB{extensao}")

    valor_00_campo_02 = None
    valor_00_campo_03 = None
    valor_00_campo_16 = None

    with open(caminho_entrada, "r", encoding="latin-1") as entrada, \
         open(caminho_saida, "w", encoding="latin-1") as saida:

        for linha in entrada:
            linha = linha.rstrip("\n")

            if not linha:
                continue

            campos = linha.split(";")
            tipo_registro = campos[0].strip()

            # =========================================
            # EXCLUSÃO DEFINITIVA DE LINHA 00 COM PEDÁGIO
            # =========================================
            if tipo_registro == "00" and contem_pedagio(linha):
                continue

            # -------------------------
            # LINHA 00 (captura dados)
            # -------------------------
            if tipo_registro == "00":
                if len(campos) >= 16:
                    valor_00_campo_02 = campos[1]
                    valor_00_campo_03 = campos[2]
                    valor_00_campo_16 = campos[15]

            # -------------------------
            # LINHA 06 (transformações)
            # -------------------------
            elif tipo_registro == "06":
                if valor_00_campo_02 and valor_00_campo_03:
                    # Campo 03 ← campo 02 da linha 00
                    campos.insert(2, valor_00_campo_02)

                    # Campo 04 ← campo 03 da linha 00
                    campos[3] = valor_00_campo_03

                    # Novo campo 19 vazio ""
                    campos.insert(18, '""')

                    # Remove campo 23
                    if len(campos) > 22:
                        del campos[22]

            saida.write(";".join(campos) + "\n")

    print("Processamento concluído com sucesso.")
    print(f"Arquivo gerado: {caminho_saida}")


def main():
    if len(sys.argv) < 2:
        print("Uso: python atuaToJb.py <arquivo_entrada>")
        sys.exit(1)

    caminho_entrada = sys.argv[1]

    if not os.path.isfile(caminho_entrada):
        print("Arquivo não encontrado:", caminho_entrada)
        sys.exit(1)

    processar_arquivo(caminho_entrada)


if __name__ == "__main__":
    main()

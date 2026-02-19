# ================================
# Importações
# ================================
import pandas as pd
import os
import re
import csv

# ================================
# Pastas para entrada e saída
# ================================
INPUT_FOLDER = "../input/"
OUTPUT_FOLDER = "../output/"

# ================================
# Converte letras de coluna (A, B, ..., AA) para índice
# ================================
def col_letter_to_index(col):
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

# ================================
# Remove tudo que não for número
# ================================
def somente_numeros(valor):
    return ''.join(re.findall(r"\d", str(valor)))

# ================================
# Coloca aspas conforme padrão
# ================================
def colocar_aspas(valor):
    if valor is None or str(valor).strip() == "":
        return "\"\""
    return f"\"{valor}\""

# ================================
# Verifica se o valor é ISENTO
# ================================
def tratar_campo04(valor):
    valor_str = str(valor).strip()

    if valor_str == "":
        return "ISENTO"

    numeros = re.findall(r"\d", valor_str)

    if numeros:
        return ''.join(numeros)

    return "ISENTO"


# ================================
# Processamento do arquivo Excel
# ================================
def processar_arquivo(caminho_excel=None):

    # -----------------------------
    # Define arquivo de entrada
    # -----------------------------
    if caminho_excel:
        arquivo_excel = caminho_excel
    else:
        if not os.path.isdir(INPUT_FOLDER):
            print(f"Pasta de input não encontrada: {INPUT_FOLDER}")
            return

        arquivo_excel = None
        for file in os.listdir(INPUT_FOLDER):
            if file.lower().endswith((".xls", ".xlsx")):
                arquivo_excel = os.path.join(INPUT_FOLDER, file)
                break

        if not arquivo_excel:
            print("Nenhum arquivo XLS/XLSX encontrado na pasta input/")
            return

    # -----------------------------
    # Leitura do Excel
    # Pula 2 linhas de cabeçalho
    # -----------------------------
    df = pd.read_excel(
        arquivo_excel,
        dtype=str,
        header=None,
        keep_default_na=False,
        skiprows=2
    )

    # -----------------------------
    # Índices das colunas usadas
    # -----------------------------
    idx_A  = col_letter_to_index("A")
    idx_B  = col_letter_to_index("B")
    idx_C  = col_letter_to_index("C")
    idx_R  = col_letter_to_index("R")
    idx_W  = col_letter_to_index("W")
    idx_X  = col_letter_to_index("X")
    idx_Y  = col_letter_to_index("Y")
    idx_AB = col_letter_to_index("AB")
    idx_AC = col_letter_to_index("AC")
    idx_AF = col_letter_to_index("AF")

    # -----------------------------
    # Mapeamento dos campos
    # -----------------------------
    col1  = df[idx_A]
    col2  = df[idx_B]
    col3  = df[idx_R].apply(somente_numeros)
    col4  = df[idx_C].apply(tratar_campo04)
    col5  = df[idx_X]
    col6  = ""
    col7  = "F"
    col8  = ""
    col9  = df[idx_Y]
    col10 = df[idx_W]
    col11 = ""
    col12 = df[idx_AF]
    col13 = "1"
    col14 = ""
    col15 = ""
    col16 = ""
    col17 = ""
    col18 = df[idx_AB]
    col19 = df[idx_AC]
    col20 = ""
    col21 = ""
    col22 = ""
    col23 = ""
    col24 = ""
    col25 = ""
    col26 = ""   # email
    col27 = ""
    col28 = ""
    col29 = "00000"
    col30 = ""
    col31 = ""
    col32 = ""

    # -----------------------------
    # DataFrame final
    # -----------------------------
    df_final = pd.DataFrame({
        1:  col1,  2:  col2,  3:  col3,  4:  col4,
        5:  col5,  6:  col6,  7:  col7,  8:  col8,
        9:  col9,  10: col10, 11: col11, 12: col12,
        13: col13, 14: col14, 15: col15, 16: col16,
        17: col17, 18: col18, 19: col19, 20: col20,
        21: col21, 22: col22, 23: col23, 24: col24,
        25: col25, 26: col26, 27: col27, 28: col28,
        29: col29, 30: col30, 31: col31, 32: col32
    })

    # -----------------------------
    # Aplica aspas em TODOS os campos
    # -----------------------------
    for col in df_final.columns:
        df_final[col] = df_final[col].apply(colocar_aspas)

    # -----------------------------
    # Geração do TXT
    # -----------------------------
    pasta_saida = os.path.dirname(arquivo_excel)
    nome_base = os.path.splitext(os.path.basename(arquivo_excel))[0]
    output_path = os.path.join(pasta_saida, f"{nome_base}_JB.txt")

    df_final.to_csv(
        output_path,
        sep=";",
        index=False,
        header=False,
        quoting=csv.QUOTE_NONE,
        escapechar='\\'
    )

    print(f"Arquivo gerado com sucesso: {output_path}")

if __name__ == "__main__":
    processar_arquivo()

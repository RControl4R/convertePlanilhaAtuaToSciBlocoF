#Importações de bibliotecas

from numpy import astype
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
# Converte letras de coluna (A, B, ..., AA) para índice numérico
# Ex: A=0, B=1, Z=25, AA=26
# ================================
def col_letter_to_index(col):
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

# ================================
# Formata datas para AAAAMMDD (Formato necessário para importação dentro do sistema SCI)
# Aceita yyyy-mm-dd, dd/mm/yyyy e número serial do Excel
# ================================
def formatar_data(valor):
    if valor is None or str(valor).strip() == "":
        return ""
    valor = str(valor).strip()
    
    # Formato yyyy-mm-dd
    if "-" in valor and len(valor.split("-")[0]) == 4:
        try:
            ano, mes, dia = valor.split(" ")[0].split("-")
            return f"{ano}{mes}{dia}"
        except: 
            pass

    # Formato dd/mm/yyyy
    if "/" in valor:
        try:
            dia, mes, ano = valor.split("/")
            return f"{ano}{mes}{dia}"
        except: 
            pass

    # Número serial do Excel
    try:
        numero = float(valor)
        data = pd.to_datetime(numero, unit="D", origin="1899-12-30")
        return data.strftime("%Y%m%d")
    except: 
        pass

    return valor

# ==================================================
# CONVERSÃO BR → FLOAT - Necessário para evitar tratar 
# números como inteiro, o que gerava erro ao calcular 
# o imposto.
# ==================================================
def br_to_float(valor):
    if valor is None:
        return 0.0

    texto = str(valor).strip()
    if texto == "":
        return 0.0

    # Caso 1: contém vírgula → formato br
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
        try:
            return float(texto)
        except:
            return 0.0

    # Caso 2: contém ponto mas não vírgula → ponto é decimal
    if "." in texto:
        try:
            return float(texto)
        except:
            return 0.0

    # Caso 3: número inteiro
    try:
        return float(texto)
    except:
        return 0.0

# ================================
# Converte float para formato customizado
# divide por 100 e arredonda para 2 casas decimais
# ================================
def float_to_custom(valor):
    return f"{round(valor/100,2):.2f}"

# ================================
# Processamento do arquivo Excel (xls xlsx)
# ================================
def processar_arquivo():
    # Localiza arquivo
    arquivo_excel = None
    for file in os.listdir(INPUT_FOLDER):
        if file.lower().endswith((".xls", ".xlsx")):
            arquivo_excel = os.path.join(INPUT_FOLDER, file)
            break
    if not arquivo_excel:
        print("Nenhum arquivo XLS/XLSX encontrado na pasta input/")
        return

    df = pd.read_excel(arquivo_excel, dtype=str, header=None, keep_default_na=False, skiprows=2) #necessário para pular duas linhas de registro pois o arquivo vem com cabeçalho

    # Ajuste de índices
    idx_H = col_letter_to_index("H")
    idx_J = col_letter_to_index("J")
    idx_V = col_letter_to_index("V")
    idx_AR = col_letter_to_index("AR")
    idx_BC = col_letter_to_index("BC")
    idx_BD = col_letter_to_index("BD")
    idx_BS = col_letter_to_index("BS")
    idx_BW = col_letter_to_index("BW") # valor fixo "SUBCONTRATAÇÃO DE FRETE " Var colBD + Fixo " N. CF " + Var colAR

    # Campo 01
    def converter_bs(valor):
        valor = str(valor).strip().lower()
        if "pessoa física" in valor: return "7"
        if "não optante" in valor: return "5"
        if "optante" in valor: return "6"
        return ""
    col1 = df[idx_BS].apply(converter_bs)

    # Campo 02
    col2 = "0"

    # Campo 03
    col3 = df[idx_BC].apply(lambda x: ''.join(re.findall(r"\d", str(x))))

    # Campo 04
    col4 = ""

    # Campo 05
    col5 = df[idx_H].apply(formatar_data)

    # Campo 06
    col6 = df[idx_J].astype(str)

    # Campo 07
    col7 = "p"

    # Campo 08
    col8 = df[idx_V].astype(str).str.strip()

    # Campo 09
    def calc_col9(bs):
        if bs == "7": return "942"
        if bs == "6": return "941"
        if bs == "5": return "940"
        return ""
    col9 = col1.apply(calc_col9)

    # Campo 10
    col10 = "815"

    # Campo 11 
    col11 = "53"

    # Campo 12
    col12 = col8 # valor original

    # Campo 13
    def calc_col13(bs):
        if bs == "5": return 1.65
        return 1.2375
    col13 = col1.apply(calc_col13)

    # ==========================================================
    # Campo 14 — Cálculo de imposto
    # ==========================================================
    col14 = [
        float_to_custom(br_to_float(v) * float(m))
        for v, m in zip(col12, col13)
    ]

    # Campo 15
    col15 = "53"

    # Campo 16 - Valor original
    col16 = col8

    # Campo 17
    def calc_col17(bs):
        if bs == "5": return 7.60
        return 5.70
    col17 = col1.apply(calc_col17)

    # ==========================================================
    # Campo 18 — cálculo de imposto
    # ==========================================================
    col18 = [
        float_to_custom(br_to_float(v) * float(m))
        for v, m in zip(col16, col17)
    ]

    # Campo 19
    col19 = "14"

    # Campo 20
    col20 = "0"

    # Campo 21
    col21 = (
        "SUBCONTRATAÇÃO DE FRETE " 
        + df[idx_BD].astype(str)
        + " N. CF "
        + df[idx_AR].astype(str)
    )

    # Colunas 22–54 vazias, exceto col 54 = col9
    extras = {i: "" for i in range(22, 54)}
    
    # Campo 54 - conta
    extras[54] = col9

    colunas_com_aspas = [2,3,4,5,6,7,8,9,10,11,15,19,21]

    df_final = pd.DataFrame({
        1: col1, 2: col2, 3: col3, 4: col4, 5: col5,
        6: col6, 7: col7, 8: col8, 9: col9, 10: col10,
        11: col11, 12: col12, 13: col13, 14: col14,
        15: col15, 16: col16, 17: col17, 18: col18,
        19: col19, 20: col20, 21: col21,
        **extras
    })

    def colocar_aspas(x):
        if x is None or x == "":
            return "\"\""
        return f"\"{x}\""

    for col in colunas_com_aspas:
        df_final[col] = df_final[col].apply(colocar_aspas)

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    nome_base = os.path.splitext(os.path.basename(arquivo_excel))[0]
    output_path = f"{OUTPUT_FOLDER}/{nome_base}_formatado.txt"

    df_final.to_csv(output_path, sep=",", index=False, header=False,
                    quoting=csv.QUOTE_NONE, escapechar='\\')

    print(f"Arquivo gerado com sucesso: {output_path}")

if __name__ == "__main__":
    processar_arquivo()

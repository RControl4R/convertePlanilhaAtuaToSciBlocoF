import pandas as pd
import os
import re

INPUT_FOLDER = "input/"
OUTPUT_FOLDER = "output/"

# Mapeamento de colunas do arquivo original
def col_letter_to_index(col):
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

# Converte data
def formatar_data(valor):
    if valor is None or str(valor).strip() == "":
        return ""

    valor = str(valor).strip()

    if "-" in valor and len(valor.split("-")[0]) == 4:
        try:
            somente_data = valor.split(" ")[0]
            ano, mes, dia = somente_data.split("-")
            return f"{dia}{mes}{ano}"
        except:
            pass

    if "/" in valor:
        try:
            dia, mes, ano = valor.split("/")
            return f"{dia}{mes}{ano}"
        except:
            pass

    try:
        numero = float(valor)
        data = pd.to_datetime(numero, unit="D", origin="1899-12-30")
        return data.strftime("%d%m%Y")
    except:
        pass

    return valor

# Converte número brasileiro → float
def br_to_float(valor):
    if valor is None:
        return 0.0
    texto = str(valor).strip()
    if texto == "":
        return 0.0

    texto = texto.replace(".", "")
    texto = texto.replace(",", ".")

    try:
        return float(texto)
    except:
        return 0.0

# Formata valores das colunas 14 e 18: últimos 4 dígitos como decimais
def formatar_colunas_valor(valor):
    try:
        num = float(valor)
        # Ajusta dividindo por 100 para pegar os 2 últimos dígitos como decimais
        num_ajustado = round(num / 100, 2)
        return f"{num_ajustado:.2f}".replace(",", ".")
    except:
        return valor

def processar_arquivo():
    arquivo_excel = None
    for file in os.listdir(INPUT_FOLDER):
        if file.lower().endswith((".xls", ".xlsx")):
            arquivo_excel = os.path.join(INPUT_FOLDER, file)
            break

    if not arquivo_excel:
        print("Nenhum arquivo XLS/XLSX encontrado na pasta input/")
        return

    df = pd.read_excel(
        arquivo_excel,
        dtype=str,
        header=None,
        keep_default_na=False
    )

    idx_H = col_letter_to_index("H")
    idx_J = col_letter_to_index("J")
    idx_V = col_letter_to_index("V")
    idx_BC = col_letter_to_index("BC")
    idx_BS = col_letter_to_index("BS")
    idx_BW = col_letter_to_index("BW")

    # Coluna 1
    def converter_bs(valor):
        valor = str(valor).strip().lower()
        if "pessoa física" in valor:
            return "7"
        if "não optante" in valor:
            return "5"
        if "optante" in valor:
            return "6"
        return ""

    col1 = df[idx_BS].apply(converter_bs)

    # Coluna 2
    col2 = "0"

    # Coluna 3
    col3 = df[idx_BC].apply(lambda x: ''.join(re.findall(r"\d", str(x))))

    # Coluna 4
    col4 = ""

    # Coluna 5
    col5 = df[idx_H].apply(formatar_data)

    # Coluna 6
    col6 = df[idx_J].astype(str)

    # Coluna 7
    col7 = "p"

    # Coluna 8
    col8 = df[idx_V].astype(str).str.strip()

    # Coluna 9
    def calc_col9(bs):
        if bs == "7": return "942"
        if bs == "6": return "941"
        if bs == "5": return "940"
        return ""

    col9 = col1.apply(calc_col9)

    # Colunas 10 e 11
    col10 = "815"
    col11 = "53"

    col12 = col8

    # Coluna 13
    def calc_col13(bs):
        if bs == "5":
            return 1.65
        return 1.2375

    col13 = col1.apply(calc_col13)

    # Coluna 14
    col14 = [
        formatar_colunas_valor(br_to_float(v) * float(m))
        for v, m in zip(col12, col13)
    ]

    # Coluna 15
    col15 = "53"

    # Coluna 16
    col16 = col8

    # Coluna 17
    def calc_col17(bs):
        if bs == "5": return 7.60
        return 5.70

    col17 = col1.apply(calc_col17)

    # Coluna 18
    col18 = [
        formatar_colunas_valor(br_to_float(v) * float(m))
        for v, m in zip(col16, col17)
    ]

    # Coluna 19
    col19 = "14"

    # Coluna 20
    col20 = df[idx_BW].astype(str)

    # Montar DataFrame final
    df_final = pd.DataFrame({
        1: col1,
        2: col2,
        3: col3,
        4: col4,
        5: col5,
        6: col6,
        7: col7,
        8: col8,
        9: col9,
        10: col10,
        11: col11,
        12: col12,
        13: col13,
        14: col14,
        15: col15,
        16: col16,
        17: col17,
        18: col18,
        19: col19,
        20: col20
    })

    # Gerar CSV separado por ponto e vírgula
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    nome_base = os.path.splitext(os.path.basename(arquivo_excel))[0]
    output_path = f"{OUTPUT_FOLDER}/{nome_base}_formatado.csv"

    df_final.to_csv(output_path, sep=";", index=False, header=False, encoding="utf-8-sig")

    print(f"Arquivo gerado com sucesso: {output_path}")

if __name__ == "__main__":
    processar_arquivo()

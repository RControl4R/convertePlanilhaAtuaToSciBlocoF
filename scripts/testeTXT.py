import pandas as pd
import os
import re
import csv

# ================================
# Pastas de entrada e saída
# ================================
INPUT_FOLDER = "input/"
OUTPUT_FOLDER = "output/"

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
# Formata datas para AAAAMMDD
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

# ================================
# Converte valor brasileiro (com ponto e vírgula) para float
# Ex: '4.637,45' → 4637.45
# ================================
def br_to_float(valor):
    if valor is None: 
        return 0.0
    texto = str(valor).strip()
    if texto == "": 
        return 0.0
    texto = texto.replace(".", "").replace(",", ".")
    try: 
        return float(texto)
    except: 
        return 0.0

# ================================
# Converte float para formato customizado, corrigindo escala
# Divide por 100 e arredonda para 2 casas decimais
# ================================
def float_to_custom(valor):
    return f"{round(valor/100,2):.2f}"  # divide por 100 para corrigir escala

# ================================
# Função principal de processamento do arquivo Excel
# ================================
def processar_arquivo():
    # ===================================
    # Localiza o primeiro arquivo XLS/XLSX na pasta de input
    # ===================================
    arquivo_excel = None
    for file in os.listdir(INPUT_FOLDER):
        if file.lower().endswith((".xls", ".xlsx")):
            arquivo_excel = os.path.join(INPUT_FOLDER, file)
            break
    if not arquivo_excel:
        print("Nenhum arquivo XLS/XLSX encontrado na pasta input/")
        return

    # ===================================
    # Lê o arquivo Excel como strings
    # ===================================
    df = pd.read_excel(arquivo_excel, dtype=str, header=None, keep_default_na=False)

    # ===================================
    # Define índices das colunas necessárias
    # ===================================
    idx_H = col_letter_to_index("H")
    idx_J = col_letter_to_index("J")
    idx_V = col_letter_to_index("V")
    idx_BC = col_letter_to_index("BC")
    idx_BS = col_letter_to_index("BS")
    idx_BW = col_letter_to_index("BW")

    # =======================
    # Coluna 1 - BS convertida
    # 7 → Pessoa Física
    # 6 → PJ Optante
    # 5 → PJ Não Optante
    # =======================
    def converter_bs(valor):
        valor = str(valor).strip().lower()
        if "pessoa física" in valor: return "7"
        if "não optante" in valor: return "5"
        if "optante" in valor: return "6"
        return ""
    col1 = df[idx_BS].apply(converter_bs)

    # =======================
    # Coluna 2 - Valor fixo "0"
    # =======================
    col2 = "0"

    # =======================
    # Coluna 3 - Extrai somente dígitos da coluna BC
    # =======================
    col3 = df[idx_BC].apply(lambda x: ''.join(re.findall(r"\d", str(x))))

    # =======================
    # Coluna 4 - Vazia
    # =======================
    col4 = ""

    # =======================
    # Coluna 5 - Data formatada para AAAAMMDD
    # =======================
    col5 = df[idx_H].apply(formatar_data)

    # =======================
    # Coluna 6 - J convertido para string
    # =======================
    col6 = df[idx_J].astype(str)

    # =======================
    # Coluna 7 - Valor fixo "p"
    # =======================
    col7 = "p"

    # =======================
    # Coluna 8 - V original
    # =======================
    col8 = df[idx_V].astype(str).str.strip()

    # =======================
    # Coluna 9 - Códigos condicionais baseados na coluna 1 (BS)
    # 7 → 942
    # 6 → 941
    # 5 → 940
    # =======================
    def calc_col9(bs):
        if bs == "7": return "942"
        if bs == "6": return "941"
        if bs == "5": return "940"
        return ""
    col9 = col1.apply(calc_col9)

    # =======================
    # Colunas 10 e 11 - valores fixos
    # =======================
    col10 = "815"
    col11 = "53"

    # =======================
    # Coluna 12 - V novamente
    # =======================
    col12 = col8

    # =======================
    # Coluna 13 - Multiplicador condicional
    # 5 → 1.65
    # Outro → 1.2375
    # =======================
    def calc_col13(bs):
        if bs == "5": return 1.65
        return 1.2375
    col13 = col1.apply(calc_col13)

    # =======================
    # Coluna 14 - Produto col12*col13, formatado com float_to_custom
    # =======================
    col14 = [float_to_custom(br_to_float(v) * float(m)/100) for v, m in zip(col12, col13)]

    # =======================
    # Coluna 15 - valor fixo "53"
    # =======================
    col15 = "53"

    # =======================
    # Coluna 16 - V novamente
    # =======================
    col16 = col8

    # =======================
    # Coluna 17 - Multiplicador condicional
    # 5 → 7.60
    # Outro → 5.70
    # =======================
    def calc_col17(bs):
        if bs == "5": return 7.60
        return 5.70
    col17 = col1.apply(calc_col17)

    # =======================
    # Coluna 18 - Produto col16*col17, formatado com float_to_custom
    # =======================
    col18 = [float_to_custom(br_to_float(v) * float(m)/100) for v, m in zip(col16, col17)]

    # =======================
    # Coluna 19 - valor fixo "14"
    # =======================
    col19 = "14"

    # =======================
    # Coluna 20 - nova coluna fixa 0
    # =======================
    col20 = "0"

    # =======================
    # Coluna 21 - BW convertido para string
    # =======================
    col21 = df[idx_BW].astype(str)

    # =======================
    # Cria colunas vazias adicionais até 54
    # =======================
    colunas_vazias = {i: "" for i in range(22,55)}

    # =======================
    # Lista de colunas que devem receber aspas na exportação
    # =======================
    colunas_com_aspas = [2,3,4,5,6,7,8,9,10,11,15,19,21]  # atualizada para nova coluna 21

    # =======================
    # Monta o DataFrame final com todas as colunas
    # =======================
    df_final = pd.DataFrame({
        1: col1, 2: col2, 3: col3, 4: col4, 5: col5,
        6: col6, 7: col7, 8: col8, 9: col9, 10: col10,
        11: col11, 12: col12, 13: col13, 14: col14,
        15: col15, 16: col16, 17: col17, 18: col18,
        19: col19, 20: col20, 21: col21, 22:"",23:"",24:"",25:"",26:"",27:"",28:"",29:"",30:"",
        31:"",32:"",33:"",34:"",35:"",36:"",37:"",38:"",39:"",40:"",
        41:"",42:"",43:"",44:"",45:"",46:"",47:"",48:"",49:"",50:"",
        51:"",52:"",53:"", 54: col9
    })

    # =======================
    # Função para colocar aspas em campos específicos
    # =======================
    def colocar_aspas(x):
        if x is None or x == "":
            return "\"\""
        return f"\"{x}\""

    # =======================
    # Aplica aspas apenas nas colunas definidas
    # =======================
    for col in colunas_com_aspas:
        df_final[col] = df_final[col].apply(colocar_aspas)

    # =======================
    # Cria pasta de saída se não existir
    # =======================
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    nome_base = os.path.splitext(os.path.basename(arquivo_excel))[0]
    output_path = f"{OUTPUT_FOLDER}/{nome_base}_formatado.txt"

    # =======================
    # Exporta DataFrame para TXT separado por vírgula
    # Sem index e sem cabeçalho, evitando aspas automáticas do pandas
    # =======================
    df_final.to_csv(output_path, sep=",", index=False, header=False, quoting=csv.QUOTE_NONE, escapechar='\\')

    print(f"Arquivo gerado com sucesso: {output_path}")

# =======================
# Executa a função principal
# =======================
if __name__ == "__main__":
    processar_arquivo()

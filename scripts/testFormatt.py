import pandas as pd # biblioteca para manipulação de dados excel 
import os # usado para navegar pelos arquivos e pastas

# Pastas de entrada/saída
INPUT_FOLDER = "input/"
OUTPUT_FOLDER = "output/"
OUTPUT_FILE = "arquivo_formatado.csv"

# Colunas que você deseja manter (por letra do Excel)
LETTERS = ["H", "T", "V", "X", "Y", "AR", "BC", "BS", "BT", "BW"]

# Função para converter letras de coluna do Excel para índice baseado em 0 
# Essa função é usada para converter as letras das colunas do Excel para índices numéricos que podem ser usados para selecionar as colunas desejadas.
def col_letter_to_index(col):
    """ Converte letra de coluna do Excel para índice baseado em 0 """
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

def processar_arquivo():
    # Procurar arquivo Excel na pasta input
    arquivo_excel = None
    for file in os.listdir(INPUT_FOLDER):
        if file.endswith(".xls") or file.endswith(".xlsx"):
            arquivo_excel = os.path.join(INPUT_FOLDER, file)
            break
    
    if not arquivo_excel:
        print("Nenhum arquivo XLS/XLSX encontrado na pasta 'input/'.")
        return

    print(f"Lendo arquivo: {arquivo_excel}")

    # Ler Excel original
    df = pd.read_excel(arquivo_excel)

    # Converter letras para nomes reais das colunas
    colunas_selecionadas = []
    for letra in LETTERS:
        idx = col_letter_to_index(letra)
        if idx < len(df.columns):
            colunas_selecionadas.append(df.columns[idx])

    # Filtrar DataFrame
    df_final = df[colunas_selecionadas]

    # Criar pasta de saída se não existir
    os

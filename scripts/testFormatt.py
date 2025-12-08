import pandas as pd  # biblioteca para manipulação de dados excel 
import os  # biblioteca para navegar pelos arquivos e pastas

# Pastas de entrada/saída
INPUT_FOLDER = "input/"
OUTPUT_FOLDER = "output/"

# Colunas desejadas para manter em novo arquivo
LETTERS = ["H", "T", "V", "X", "Y", "AR", "BC", "BS", "BT", "BW"]

# Função para converter letra de coluna do Excel para índice baseado em 0
def col_letter_to_index(col):
    """Converte letra de coluna do Excel para índice baseado em 0"""
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1


# 🔹 Função para formatar datas no padrão DDMMAAAA
def formatar_data(valor):
    if valor is None or str(valor).strip() == "":
        return ""

    valor = str(valor).strip()

    # Se vier como "2025-10-29 14:45:06" ou "2025-10-29"
    if "-" in valor and len(valor.split("-")[0]) == 4:
        try:
            somente_data = valor.split(" ")[0]
            ano, mes, dia = somente_data.split("-")
            return f"{dia}{mes}{ano}"
        except:
            pass

    # Se vier como "29/10/2025"
    if "/" in valor:
        try:
            dia, mes, ano = valor.split("/")
            return f"{dia}{mes}{ano}"
        except:
            pass

    # Se vier como número serial do Excel
    try:
        numero = float(valor)
        data = pd.to_datetime(numero, unit="D", origin="1899-12-30")
        return data.strftime("%d%m%Y")
    except:
        pass

    return valor


# Função para processar o arquivo Excel
def processar_arquivo():

    # Procurar arquivo Excel
    arquivo_excel = None
    for file in os.listdir(INPUT_FOLDER):
        if file.lower().endswith((".xls", ".xlsx")):
            arquivo_excel = os.path.join(INPUT_FOLDER, file)
            break

    if not arquivo_excel:
        print("Nenhum arquivo XLS/XLSX encontrado na pasta 'input/'.")
        return

    print(f"Lendo arquivo: {arquivo_excel}")

    # ▶ PRIMEIRA LEITURA: descobrir quantas colunas existem
    df_tmp = pd.read_excel(arquivo_excel, nrows=1, header=None)

    # Criar converters apenas para as colunas realmente existentes
    num_cols = len(df_tmp.columns)
    converters = {i: (lambda x: str(x)) for i in range(num_cols)}

    # ▶ SEGUNDA LEITURA: agora sim, ler tudo como texto e sem cabeçalho
    df = pd.read_excel(
        arquivo_excel,
        converters=converters,
        header=None  # <<< ESSENCIAL: não tratar primeira linha como cabeçalho
    )

    # Mapear nomes reais das colunas desejadas (agora indexadas por número)
    colunas_selecionadas = []
    for letra in LETTERS:
        idx = col_letter_to_index(letra)
        if idx < len(df.columns):
            colunas_selecionadas.append(idx)

    # Filtrar DataFrame
    df_final = df[colunas_selecionadas].copy()

    # -----------------------------
    # FORMATAR COLUNA H (DATA)
    # -----------------------------
    col_H = df_final.columns[0]
    df_final[col_H] = df_final[col_H].apply(formatar_data)

    # -----------------------------
    # NÃO formatar coluna V
    # Manter exatamente como vem do arquivo original
    # -----------------------------
    col_V = df_final.columns[2]
    df_final[col_V] = df_final[col_V].astype(str).str.strip()

    # -----------------------------
    # FORMATAR COLUNA BC (somente números)
    # -----------------------------
    col_BC = df_final.columns[6]  # BC é a 7ª coluna selecionada
    df_final[col_BC] = df_final[col_BC].apply(lambda x: ''.join(filter(str.isdigit, str(x))))

    # -----------------------------
    # FORMATAR COLUNA BS (categoria -> código)
    # -----------------------------
    col_BS = df_final.columns[7]  # BS é a 8ª coluna selecionada

    def converter_bs(valor):
        valor = str(valor).strip().lower()

        if "pessoa física" in valor:
            return "7"
        if "optante" in valor:  # Pessoa Jurídica - Optante p/ Simples
            return "6"
        if "não optante" in valor or "nao optante" in valor:
            return "5"

        return ""  # caso não reconheça a categoria

    df_final[col_BS] = df_final[col_BS].apply(converter_bs)

    # Criar pasta de saída se não existir
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # Criar nome do arquivo de saída com base no arquivo original
    nome_base = os.path.splitext(os.path.basename(arquivo_excel))[0]
    nome_saida = f"{nome_base}_formatado.csv"
    output_path = os.path.join(OUTPUT_FOLDER, nome_saida)

    # Gerar CSV com delimitador ";" e sem cabeçalho
    df_final.to_csv(
        output_path,
        sep=";",
        index=False,
        header=False,
        encoding="utf-8-sig"
    )

    print(f"Arquivo gerado com sucesso em: {output_path}")


if __name__ == "__main__":
    processar_arquivo()

import pandas as pd
import os
from datetime import datetime
import warnings

# Desabilita warnings de estilos do openpyxl (opcional)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Caminhos das pastas de origem e destino
origem = r'C:\\_arquivos_acoes\\processar\\'
destino = r'C:\\_arquivos_acoes\\processado\\'

# Criar pastas se não existir
if not os.path.exists(origem):
    os.makedirs(origem)
if not os.path.exists(destino):
    os.makedirs(destino)

# Marca o início do processamento
inicio_processamento = datetime.now()
arquivos = os.listdir(origem)

# Filtrar apenas arquivos Excel válidos (ignorando temporários)
lista_arquivo = [os.path.join(origem, nome_arquivo) for nome_arquivo in arquivos if not nome_arquivo.startswith('~$') and nome_arquivo.endswith('.xlsx')]

# Lê todas as abas do arquivo como um dicionário de DataFrames
df_dict = pd.read_excel(lista_arquivo[0], sheet_name=None)

# Iterar sobre cada aba e aplicar dropna() para remover linhas nulas
for nome_aba, df in df_dict.items():
    # Remove linhas onde a coluna 'Quantidade' ou outras colunas chave tenham o valor 'Total' ou valores nulos
    df_dict[nome_aba] = df[~df.isin(['Total']).any(axis=1)]  # Filtra as linhas que têm o valor 'Total'
    
    # Remove a última linha, se ela for nula ou vazia
    df_dict[nome_aba] = df_dict[nome_aba].dropna(how='all')  # Remove linhas totalmente vazias
    
    tipos_dados = {
        'Conta': str,
        'CNPJ da Empresa': str,
        'Quantidade': 'Int64',
        'Quantidade Disponível': 'Int64',
        'Quantidade Indisponível': 'Int64',
        'Valor Atualizado': float,
        'CNPJ do Fundo': str,
        'Valor líquido': float,
        'Período (Inicial)': 'datetime64',
        'Período (Final)': 'datetime64',
        'Preço Médio (Venda)': float
    }

    # Percorre as colunas e aplica a conversão de tipos, se a coluna existir no DataFrame
    for coluna, tipo in tipos_dados.items():
        if coluna in df.columns:
            # Se for uma coluna numérica, converte usando pd.to_numeric para lidar com NaN
            if tipo == 'Int64':
                df[coluna] = pd.to_numeric(df[coluna], errors='coerce').astype('Int64')
            else:
                df[coluna] = df[coluna].astype(tipo)  # Aplica a conversão de tipo especificado

    print(f'ABA: {nome_aba}\n{df.dtypes}\n---')

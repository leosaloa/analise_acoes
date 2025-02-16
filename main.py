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

# Início do processamento
inicio_processamento = datetime.now()
arquivos = os.listdir(origem)

# Filtrar apenas arquivos Excel válidos (ignorando temporários)
lista_arquivo = [os.path.join(origem, nome_arquivo) for nome_arquivo in arquivos if not nome_arquivo.startswith('~$') and nome_arquivo.endswith('.xlsx')]

# Lê todas as abas do arquivo como um dicionário de DataFrames
df_dict = pd.read_excel(lista_arquivo[0], sheet_name=None)

# Iterar sobre cada aba e aplicar dropna() para remover linhas nulas
nome_abas = []
conteudo_abas = []
for nome_aba, df in df_dict.items():
    df_dict[nome_aba] = df.dropna(subset=['Instituição'])
    nome_abas.append(nome_aba)
    conteudo_abas.append(df_dict[nome_aba])

tipos_dados_geral = {
    'Produto': str,
    'Instituição': str,
    'Conta': 'Int64',
    'Código de Negociação': str,
    'CNPJ da Empresa': str,
    'CNPJ do Fundo': str,
    'Código ISIN / Distribuição': str,
    'Tipo': str,
    'Escriturador': str,
    'Administrador': str,
    'Quantidade': 'int64',
    'Quantidade Disponível': 'int64',
    'Quantidade Indisponível': 'int64',
    'Motivo': str,
    'Preço de Fechamento': float,
    'Valor Atualizado': float,
    'Pagamento': 'Int64',
    'Tipo de Evento': str,
    'Preço unitário': float,
    'Valor líquido': float,
    'Período (Inicial)': 'datetime64',
    'Período (Final)': 'datetime64',
    'Quantidade (Compra)': 'int64',
    'Quantidade (Venda)': 'int64',
    'Quantidade (Líquida)': 'int64',
    'Preço Médio (Compra)': float,
    'Preço Médio (Venda)': float
}

# Remover '-' de todas as abas e colunas de texto
for nome_aba, df in df_dict.items():
    for col in df.select_dtypes(include=['object', 'string']):
        df_dict[nome_aba][col] = df[col].map(lambda x: x.replace('-', '') if isinstance(x, str) else x)

# Caminho do arquivo exportado
arquivo_saida = os.path.join(destino, 'arquivo_tratado.xlsx')

# Criando um ExcelWriter para salvar múltiplas abas
with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    for nome_aba, df in df_dict.items():
        df.to_excel(writer, sheet_name=nome_aba, index=False)  # Salva cada aba no arquivo Excel
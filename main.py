import pandas as pd
import os
from datetime import datetime
import warnings

# CRIAR FUNÇÕES PARA MELHOR MANUTENÇÃO

# Desabilita warnings de estilos do openpyxl (opcional)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Caminhos das pastas de origem e destino
origem_c = r'C:\\_arquivos_acoes\\'
origem = r'C:\\_arquivos_acoes\\processar\\'
destino = r'C:\\_arquivos_acoes\\processado\\'

# Criar pastas se não existir
os.makedirs(origem_c, exist_ok=True)
os.makedirs(origem, exist_ok=True)
os.makedirs(destino, exist_ok=True)

# Início do processamento
inicio_processamento = datetime.now()
arquivos = os.listdir(origem)

# Filtrar apenas arquivos Excel válidos (ignorando temporários)
lista_arquivos = [os.path.join(origem, nome) for nome in arquivos if nome.endswith('.xlsx') and not nome.startswith('~$')]

# Dicionário de tipos
tipos_dados_geral = {
    'produto': str,
    'instituição': str,
    'conta': 'Int64',
    'código de negociação': str,
    'cnpj da empresa': str,
    'cnpj do fundo': str,
    'código isin / distribuição': str,
    'tipo': str,
    'escriturador': str,
    'administrador': str,
    'quantidade': 'Int64',
    'quantidade disponível': 'Int64',
    'quantidade indisponível': 'Int64',
    'motivo': str,
    'preço de fechamento': float,
    'valor atualizado': float,
    'pagamento': 'datetime64',
    'tipo de Evento': str,
    'preço unitário': float,
    'valor líquido': float,
    'período (inicial)': 'datetime64',
    'período (final)': 'datetime64',
    'quantidade (compra)': 'Int64',
    'quantidade (venda)': 'Int64',
    'quantidade (líquida)': 'Int64',
    'preço médio (compra)': float,
    'preço médio (venda)': float,
    'data relatório': 'datetime64'
}

col_remove_hifen = ['quantidade', 'quantidade disponível', 'quantidade indisponível', 'quantidade (compra)',
                     'quantidade (venda)', 'quantidade (líquida)', 'preço de fechamento', 
                     'valor atualizado', 'preço unitário', 'valor líquido', 
                     'preço médio (compra)', 'preço médio (venda)', 'motivo', 'período (final)']

# Dicionário final com todas as abas dos arquivos
df_dict_final = {}

# Processar cada arquivo
for caminho_arquivo in lista_arquivos:
    nome_arquivo = os.path.basename(caminho_arquivo)
    try:
        df_dict = pd.read_excel(caminho_arquivo, sheet_name=None)
    except ValueError as e:
        print(f'⚠️ Arquivo ignorado (sem abas) {nome_arquivo}: {e}')
        continue
    except Exception as e:
        print(f'⚠️ Erro ao ler o arquivo {nome_arquivo}: {e}')
        continue

    mes_ano = nome_arquivo[29:-5].split('-')
    ano_relatorio = mes_ano[0]
    if mes_ano[1] == 'marco':
        mes_relatorio = 'março'
    else:
        mes_relatorio = mes_ano[1]

    for nome_aba, df in df_dict.items():
        df.columns = df.columns.str.strip().str.lower()
        df_dict[nome_aba].loc[:, 'data relatório'] = f'{mes_relatorio}/{ano_relatorio}'
        df_dict[nome_aba] = df

    for nome_aba, df in df_dict.items():
        df = df.dropna(subset=['instituição']).copy()

        if 'produto' in df.columns:
            df.loc[:, 'código de negociação'] = df['produto'].str.split(' - ').str[0]

        for col, tipo in tipos_dados_geral.items():
            if col in df.columns:
                try:
                    if tipo in ['int64', 'Int64']:
                        df.loc[:, col] = pd.to_numeric(
                            df[col].astype(str)
                            .str.replace('R$', '', regex=False)
                            .str.strip(),
                            errors='coerce'
                        )
                        # Se quiser forçar para inteiro, tente apenas se não houver valores com decimais
                        if df[col].dropna().apply(lambda x: float(x).is_integer).all():
                            df.loc[:, col] = df[col].astype('Int64')
                        elif tipo == 'float':
                            df.loc[:, col] = pd.to_numeric(
                                df[col].astype(str)
                                .str.replace(',', '.')
                                .str.replace('R$', '', regex=False)
                                .str.strip(),
                                errors='coerce'
                            )
                        elif tipo == 'datetime64':
                            df.loc[:, col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
                        else:
                            df.loc[:, col] = df[col].astype(str)
                    # Tratamento específico para colunas com hífen
                    if col in col_remove_hifen:
                        df.loc[:, col] = pd.to_numeric(
                            df[col].astype(str)
                            .str.replace('-', '', regex=False)  # Remove hífen
                        )
                    if col in ['cnpj da empresa', 'cnpj do fundo']:
                        df.loc[:, col] = df[col].astype(str).str.replace('-', '', regex=True)
                except Exception as e:
                    pass

        # Acumular os dados
        if not df.empty:
            if nome_aba in df_dict_final:
                df_dict_final[nome_aba] = pd.concat([df_dict_final[nome_aba], df], ignore_index=True)
            else:
                df_dict_final[nome_aba] = df

# Exportar se houver dados
# extrair = input('Deseja extrair os dados? (S/N): ').strip().upper()
extrair = 'S'

if extrair == 'S':
    if df_dict_final:
        arquivo_saida = os.path.join(destino, 'arquivo_tratado_unico.xlsx')
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            for nome_aba, df in df_dict_final.items():
                df.to_excel(writer, sheet_name=nome_aba[:31], index=False)
        print(f'✅ Processamento finalizado. Arquivo salvo em: {arquivo_saida}')
    else:
        print('⚠️ Nenhum dado válido encontrado para exportar.')
else:
    pass

# Fim do processamento
fim_processamento = datetime.now()

print(f'⏱️ Tempo total de processamento: {fim_processamento - inicio_processamento}\n')

dt_rel = []
vlr_atu = []

for nome_aba, df in df_dict_final.items():
    if nome_aba in ('Posição - Ações', 'Posição - BDR', 'Posição - Fundos', 'Posição - Tesouro Direto'):
        for col in df:
            if col == 'valor atualizado':
                vlr_atu.append(df['valor atualizado'])
            if col == 'data relatório':
                dt_rel.append(df['data relatório'])
print(dt_rel, vlr_atu)

    # for col in df:
        # if nome_aba == 'Posição - Ações' and col == 'preço de fechamento':
            # preco_fechamento = sum(df[col].dropna())
            # print(preco_fechamento)
# print(df_dict_final)



# for nome_aba, df in df_dict_final.items():
    # for col in df:
        # if nome_aba == 'Posição - Ações' and col == 'preço de fechamento':
            # preco_fechamento = sum(df[col].dropna())
            # print(preco_fechamento)
# print(df_dict_final)

# SUM('Posição - Ações'[valor atualizado]) + 
# SUM('Posição - BDR'[valor atualizado]) + 
# SUM('Posição - Fundos'[valor atualizado]) + 
# SUM('Posição - Tesouro Direto'[valor atualizado])
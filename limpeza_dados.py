import pandas as pd

# Carregar o dataset
dados = pd.read_excel('dados_venda.xlsx')  # Substituir pelo nome correto do arquivo

# Exibir as primeiras linhas para inspeção inicial
print(dados.head())

# --- TRATAMENTO DE VALORES AUSENTES ---
print("Valores ausentes antes do tratamento:")
print(dados.isnull().sum())

# Preencher valores ausentes com valores padrão
dados['IDADE'].fillna(dados['IDADE'].mean(), inplace=True)
dados.fillna({'QTD_UNIDADE_FARMACOTECNICA': 0, 'SEXO': 'Desconhecido'}, inplace=True)

# --- EXCLUSÃO DE DADOS COM IDADE INVÁLIDA ---
# Filtrar apenas idades acima de 18 anos
dados = dados[dados['IDADE'] >= 18]

# --- REMOÇÃO DE DUPLICATAS ---
dados.drop_duplicates(inplace=True)

# --- PADRONIZAÇÃO DE DATAS ---
dados['MÊS_VENDA'] = pd.to_datetime(dados['MÊS_VENDA'], format='%m/%d/%Y', errors='coerce').dt.date
dados['MÊS_VENCIMENTO'] = pd.to_datetime(dados['MÊS_VENCIMENTO'], format='%m/%d/%Y', errors='coerce').dt.date

# Exibir dados após a conversão de datas
print("Dados após a conversão de datas:")
print(dados[['MÊS_VENDA', 'MÊS_VENCIMENTO']].head())

# --- RESUMO DOS DADOS TRATADOS ---
print("Resumo dos dados tratados:")
print(dados.info())

# Salvar o dataset limpo em formato .xlsx
dados.to_excel('dados_vendas_limpo.xlsx', index=False)

import pandas as pd

# Carregar o dataset (agora em formato .xlsx)
dados = pd.read_excel('dados_vendas.xlsx')

# Exibir as primeiras linhas para inspeção inicial
print(dados.head())

# --- TRATAMENTO DE VALORES AUSENTES --- 
print("Valores ausentes antes do tratamento:")
print(dados.isnull().sum())

# Tratar valores ausentes
dados['IDADE'].fillna(dados['IDADE'].mean(), inplace=True)
dados.fillna({'QTD_UNIDADE_FARMACOTECNICA': 0, 'SEXO': 'Desconhecido'}, inplace=True)

# --- EXCLUSÃO DE DADOS COM IDADE INVÁLIDA --- 
# Remover linhas onde a idade é nula, zero ou em branco
dados = dados[dados['IDADE'].notnull()]  # Remover valores nulos
dados = dados[dados['IDADE'] > 0]        # Remover idades menores ou iguais a 0
dados = dados[dados['IDADE'] >= 18]      # Incluir apenas idades de 18 anos ou mais

# --- REMOÇÃO DE DUPLICATAS ---
dados.drop_duplicates(inplace=True)

# --- PADRONIZAÇÃO DE DATAS ---
dados['MÊS_VENDA'] = pd.to_datetime(dados['MÊS_VENDA'], format='%m/%d/%Y', errors='coerce').dt.date
dados['MÊS_VENCIMENTO'] = pd.to_datetime(dados['MÊS_VENCIMENTO'], format='%m/%d/%Y', errors='coerce').dt.date

# Exibir dados após a conversão de datas
print("Dados após a conversão de datas:")
print(dados[['MÊS_VENDA', 'MÊS_VENCIMENTO']].head())

# --- PADRONIZAÇÃO DE VALORES MONETÁRIOS ---
if 'VALOR_VENDA' in dados.columns:
    dados['VALOR_VENDA'] = dados['VALOR_VENDA'].fillna(0).astype(float)

# Exibir resumo dos dados tratados
print("Resumo dos dados tratados:")
print(dados.info())

# Salvar o dataset limpo em formato .xlsx
dados.to_excel('dados_vendas_limpo.xlsx', index=False)

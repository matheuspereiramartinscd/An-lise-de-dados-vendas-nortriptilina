import pandas as pd
import unicodedata

# Função para remover acentos
def remover_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

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

# Renomear a coluna SEXO para GÊNERO
dados.rename(columns={'SEXO': 'GÊNERO'}, inplace=True)

# Substituir valores na coluna GÊNERO
dados['GÊNERO'] = dados['GÊNERO'].replace({
    1: 'Masculino',
    2: 'Feminino',
    'Desconhecido': 'Não informado'
})

# --- EXCLUSÃO DE DADOS COM IDADE INVÁLIDA ---
# Filtrar apenas idades acima de 18 anos
dados = dados[dados['IDADE'] >= 18]

# --- REMOÇÃO DE DUPLICATAS ---
dados.drop_duplicates(inplace=True)

# --- PADRONIZAÇÃO DE DATAS ---
# Excluir a coluna 'MÊS_VENCIMENTO'
dados.drop(columns=['MÊS_VENCIMENTO'], inplace=True)

# Converter 'MÊS_VENDA' para o nome do mês em português
dados['MÊS_VENDA'] = pd.to_datetime(dados['MÊS_VENDA'], format='%m/%d/%Y', errors='coerce').dt.month

# Mapeamento direto dos números dos meses para os nomes dos meses em português
meses = {
    1: 'janeiro', 2: 'fevereiro', 3: 'marco', 4: 'abril', 5: 'maio', 6: 'junho',
    7: 'julho', 8: 'agosto', 9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
}

# Substituir os números do mês pelo nome do mês em português
dados['MÊS_VENDA'] = dados['MÊS_VENDA'].map(meses)

# Remover acentos dos nomes dos meses
dados['MÊS_VENDA'] = dados['MÊS_VENDA'].apply(remover_acentos)

# Exibir dados após a conversão de datas
print("Dados após a conversão de datas:")
print(dados[['MÊS_VENDA']].head())

# --- RESUMO DOS DADOS TRATADOS ---
print("Resumo dos dados tratados:")
print(dados.info())

# Salvar o dataset limpo em formato .xlsx
dados.to_excel('dados_vendas_limpo.xlsx', index=False)

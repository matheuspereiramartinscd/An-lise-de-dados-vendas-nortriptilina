import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Carregar o dataset
dados = pd.read_excel('dados_vendas.xlsx')

# --- 1. Faturamento Total ---
# Caso tenha uma coluna de VALOR_VENDA, calcule o faturamento
# Se a coluna não existir, você pode calcular manualmente
faturamento_total = dados['VALOR_VENDA'].sum() if 'VALOR_VENDA' in dados.columns else "Coluna VALOR_VENDA não encontrada"
print(f"Faturamento Total: {faturamento_total}")

# --- 2. Quantidade Total de Produtos Vendidos ---
qtd_total_produtos = dados['QTD_UNIDADE_FARMACOTECNICA'].sum()
print(f"Quantidade Total de Produtos Vendidos: {qtd_total_produtos}")

# --- 3. Média de Vendas por Mês ---
# Convertendo 'MÊS_VENDA' para o formato de data e calculando a média mensal
dados['MÊS_VENDA'] = pd.to_datetime(dados['MÊS_VENDA'], format='%m/%d/%Y', errors='coerce')
vendas_por_mes = dados.groupby(dados['MÊS_VENDA'].dt.to_period('M'))['QTD_UNIDADE_FARMACOTECNICA'].sum()
media_vendas_por_mes = vendas_por_mes.mean()
print(f"Média de Vendas por Mês: {media_vendas_por_mes}")

# --- 4. Vendas por Sexo ---
vendas_por_sexo = dados.groupby('SEXO')['QTD_UNIDADE_FARMACOTECNICA'].sum()
print("Vendas por Sexo: ")
print(vendas_por_sexo)

# --- 5. Vendas por Estado (UF) ---
vendas_por_estado = dados.groupby('UF_VENDA')['QTD_UNIDADE_FARMACOTECNICA'].sum()
print("Vendas por Estado: ")
print(vendas_por_estado)

# --- 6. Vendas por Município ---
vendas_por_municipio = dados.groupby('MUNICIPIO_VENDA')['QTD_UNIDADE_FARMACOTECNICA'].sum()
print("Vendas por Município: ")
print(vendas_por_municipio)

# --- 7. Distribuição de Idade dos Clientes ---
sns.histplot(dados['IDADE'], kde=True)
plt.title('Distribuição de Idade dos Clientes')
plt.xlabel('Idade')
plt.ylabel('Frequência')
plt.show()

# --- 8. Vendas por Ano ---
vendas_por_ano = dados.groupby('ANO_VENDA')['QTD_UNIDADE_FARMACOTECNICA'].sum()
print("Vendas por Ano: ")
print(vendas_por_ano)

# --- 9. Tendência de Vendas ao Longo do Tempo ---
vendas_mensais = dados.groupby(dados['MÊS_VENDA'].dt.to_period('M'))['QTD_UNIDADE_FARMACOTECNICA'].sum()
sns.lineplot(x=vendas_mensais.index.astype(str), y=vendas_mensais.values)
plt.title('Tendência de Vendas por Mês')
plt.xlabel('Mês')
plt.ylabel('Quantidade Vendida')
plt.xticks(rotation=45)
plt.show()

# --- 10. Correlação entre Idade e Quantidade Vendida ---
sns.heatmap(dados[['IDADE', 'QTD_UNIDADE_FARMACOTECNICA']].corr(), annot=True, cmap='coolwarm')
plt.title('Correlação entre Idade e Quantidade Vendida')
plt.show()

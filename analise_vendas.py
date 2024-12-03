import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Carregar o dataset
dados = pd.read_excel('dados_vendas.xlsx')

# --- 1. Faturamento Total ---
faturamento_total = dados['VALOR_VENDA'].sum() if 'VALOR_VENDA' in dados.columns else "Coluna VALOR_VENDA não encontrada"
print(f"Faturamento Total: {faturamento_total}")
# Insight: Se o faturamento total for elevado, pode indicar um bom desempenho comercial, ou maior procura por medicamentos.

# --- 2. Quantidade Total de Produtos Vendidos ---
qtd_total_produtos = dados['QTD_UNIDADE_FARMACOTECNICA'].sum()
print(f"Quantidade Total de Produtos Vendidos: {qtd_total_produtos}")
# Insight: Uma quantidade de produtos vendida elevada pode indicar que o mercado está bem abastecido ou que há uma alta demanda por medicamentos.

# --- 3. Média de Vendas por Mês ---
dados['MÊS_VENDA'] = pd.to_datetime(dados['MÊS_VENDA'], format='%m/%d/%Y', errors='coerce')
vendas_por_mes = dados.groupby(dados['MÊS_VENDA'].dt.to_period('M'))['QTD_UNIDADE_FARMACOTECNICA'].sum()
media_vendas_por_mes = vendas_por_mes.mean()
print(f"Média de Vendas por Mês: {media_vendas_por_mes}")
# Insight: A média de vendas por mês mostra a tendência de consumo mensal. Se a média estiver aumentando, isso pode indicar que a demanda está crescendo.

# --- 4. Mediana de Vendas por Mês ---
mediana_vendas_por_mes = vendas_por_mes.median()
print(f"Mediana de Vendas por Mês: {mediana_vendas_por_mes}")
# Insight: A mediana é útil para entender a distribuição das vendas. Se houver uma grande diferença entre a média e a mediana, pode indicar que há outliers (como picos de vendas em alguns meses).

# --- 5. Desvio Padrão de Vendas por Mês ---
desvio_padrao_vendas_por_mes = vendas_por_mes.std()
print(f"Desvio Padrão de Vendas por Mês: {desvio_padrao_vendas_por_mes}")
# Insight: O desvio padrão ajuda a entender a volatilidade nas vendas. Se o desvio for alto, isso pode indicar que as vendas são muito variáveis de mês para mês.

# --- 6. Vendas por Sexo ---
vendas_por_sexo = dados.groupby('SEXO')['QTD_UNIDADE_FARMACOTECNICA'].sum()
print("Vendas por Sexo: ")
print(vendas_por_sexo)
# Insight: Se as vendas de um medicamento estiverem concentradas em um gênero específico, isso pode sugerir que ele é mais utilizado por aquele grupo, possivelmente devido a fatores demográficos ou médicos.

# --- 7. Vendas por Estado (UF) ---
vendas_por_estado = dados.groupby('UF_VENDA')['QTD_UNIDADE_FARMACOTECNICA'].sum()
print("Vendas por Estado: ")
print(vendas_por_estado)
# Insight: Se um estado X tiver vendas muito mais altas do que outros, isso pode indicar uma maior prevalência de uma condição tratada pelo medicamento ou uma maior conscientização sobre o tratamento naquela região.

# --- 8. Vendas por Município ---
vendas_por_municipio = dados.groupby('MUNICIPIO_VENDA')['QTD_UNIDADE_FARMACOTECNICA'].sum()
print("Vendas por Município: ")
print(vendas_por_municipio)
# Insight: Analisar municípios com grandes vendas pode revelar padrões locais de saúde, possivelmente indicando maior procura por tratamentos em certas áreas.

# --- 9. Distribuição de Idade dos Clientes ---
dados_filtrados = dados[dados['IDADE'] > 10]  # Excluir idades 0
sns.histplot(dados_filtrados['IDADE'], kde=True)
plt.title('Distribuição de Idade dos Clientes (Sem Idade 0)')
plt.xlabel('Idade')
plt.ylabel('Frequência')
plt.show()  # Exibe o gráfico
# Insight: A distribuição de idades pode indicar em quais faixas etárias o medicamento é mais utilizado. Se houver um pico em idades mais altas, pode sugerir uma maior prevalência de doenças associadas à idade, como depressão ou doenças crônicas.

# --- 10. Vendas por Ano ---
vendas_por_ano = dados.groupby('ANO_VENDA')['QTD_UNIDADE_FARMACOTECNICA'].sum()
print("Vendas por Ano: ")
print(vendas_por_ano)
# Insight: A análise de vendas por ano pode mostrar tendências de crescimento ou diminuição nas vendas. Se houver um aumento contínuo, pode ser devido a campanhas de marketing ou um aumento na necessidade do medicamento.

# --- 11. Tendência de Vendas ao Longo do Tempo ---
vendas_mensais = dados.groupby(dados['MÊS_VENDA'].dt.to_period('M'))['QTD_UNIDADE_FARMACOTECNICA'].sum()
sns.lineplot(x=vendas_mensais.index.astype(str), y=vendas_mensais.values)
plt.title('Tendência de Vendas por Mês')
plt.xlabel('Mês')
plt.ylabel('Quantidade Vendida')
plt.xticks(rotation=45)
plt.show()  # Exibe o gráfico
# Insight: Se houver uma tendência crescente nas vendas ao longo dos meses, pode sugerir que a conscientização sobre o medicamento ou a sua demanda está aumentando. Uma queda abrupta pode ser investigada para entender o que causou a diminuição.

# --- 12. Correlação entre Idade e Quantidade Vendida ---
sns.scatterplot(x=dados_filtrados['IDADE'], y=dados_filtrados['QTD_UNIDADE_FARMACOTECNICA'])
plt.title('Correlação entre Idade e Quantidade Vendida')
plt.xlabel('Idade')
plt.ylabel('Quantidade Vendida')
plt.show()  # Exibe o gráfico
# Insight: Se houver uma correlação positiva entre a idade e a quantidade vendida, isso pode indicar que pessoas mais velhas estão comprando mais do medicamento, o que pode estar relacionado ao tratamento de doenças prevalentes em faixas etárias mais altas.

# --- 13. Análise de Outliers (Vendas por Mês) ---
Q1 = vendas_por_mes.quantile(0.25)
Q3 = vendas_por_mes.quantile(0.75)
IQR = Q3 - Q1
outliers = vendas_por_mes[(vendas_por_mes < (Q1 - 1.5 * IQR)) | (vendas_por_mes > (Q3 + 1.5 * IQR))]
print("Outliers em Vendas por Mês: ")
print(outliers)
# Insight: Os outliers podem indicar meses de vendas excepcionais, que podem ter sido influenciados por campanhas promocionais, sazonalidade ou outros fatores. Investigar esses meses pode revelar padrões importantes.

# --- 14. Vendas por Tipo de Receituário ---
vendas_por_tipo_receituario = dados.groupby('TIPO_RECEITUARIO')['QTD_UNIDADE_FARMACOTECNICA'].sum()
print("Vendas por Tipo de Receituário: ")
print(vendas_por_tipo_receituario)
# Insight: A análise por tipo de receituário pode indicar se o medicamento é mais frequentemente prescrito por médicos especialistas (como psiquiatras) ou se é mais acessível via receituário simples.

# --- 15. Vendas por Princípio Ativo ---
vendas_por_principio_ativo = dados.groupby('PRINCIPIO_ATIVO')['QTD_UNIDADE_FARMACOTECNICA'].sum()
print("Vendas por Princípio Ativo: ")
print(vendas_por_principio_ativo)
# Insight: A análise de vendas por princípio ativo pode revelar quais substâncias são mais populares ou usadas para tratar uma condição específica. Se um princípio ativo específico se destacar, isso pode indicar sua eficácia ou a prevalência de uma doença tratada por ele.

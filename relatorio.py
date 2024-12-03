import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from datetime import datetime

# Função para gerar relatório em HTML
def gerar_relatorio(dados):
    relatorio_html = f"<h1>Relatório de Análise de Vendas - {datetime.now().strftime('%d/%m/%Y')}</h1>"

    # --- 1. Faturamento Total ---
    faturamento_total = dados['VALOR_VENDA'].sum() if 'VALOR_VENDA' in dados.columns else "Coluna VALOR_VENDA não encontrada"
    relatorio_html += f"<h2>Faturamento Total</h2><p>{faturamento_total}</p>"

    # --- 2. Quantidade Total de Produtos Vendidos ---
    qtd_total_produtos = dados['QTD_UNIDADE_FARMACOTECNICA'].sum()
    relatorio_html += f"<h2>Quantidade Total de Produtos Vendidos</h2><p>{qtd_total_produtos}</p>"

    # --- 3. Média de Vendas por Mês ---
    dados['MÊS_VENDA'] = pd.to_datetime(dados['MÊS_VENDA'], format='%m/%d/%Y', errors='coerce')
    vendas_por_mes = dados.groupby(dados['MÊS_VENDA'].dt.to_period('M'))['QTD_UNIDADE_FARMACOTECNICA'].sum()
    media_vendas_por_mes = vendas_por_mes.mean()
    relatorio_html += f"<h2>Média de Vendas por Mês</h2><p>{media_vendas_por_mes}</p>"

    # --- 4. Mediana de Vendas por Mês ---
    mediana_vendas_por_mes = vendas_por_mes.median()
    relatorio_html += f"<h2>Mediana de Vendas por Mês</h2><p>{mediana_vendas_por_mes}</p>"

    # --- 5. Desvio Padrão de Vendas por Mês ---
    desvio_padrao_vendas_por_mes = vendas_por_mes.std()
    relatorio_html += f"<h2>Desvio Padrão de Vendas por Mês</h2><p>{desvio_padrao_vendas_por_mes}</p>"

    # --- 6. Vendas por Sexo ---
    vendas_por_sexo = dados.groupby('SEXO')['QTD_UNIDADE_FARMACOTECNICA'].sum()
    relatorio_html += f"<h2>Vendas por Sexo</h2><p>{vendas_por_sexo}</p>"

    # --- 7. Vendas por Estado (UF) ---
    vendas_por_estado = dados.groupby('UF_VENDA')['QTD_UNIDADE_FARMACOTECNICA'].sum()
    relatorio_html += f"<h2>Vendas por Estado</h2><p>{vendas_por_estado}</p>"

    # --- 8. Vendas por Município ---
    vendas_por_municipio = dados.groupby('MUNICIPIO_VENDA')['QTD_UNIDADE_FARMACOTECNICA'].sum()
    relatorio_html += f"<h2>Vendas por Município</h2><p>{vendas_por_municipio}</p>"

    # --- 9. Distribuição de Idade dos Clientes ---
    dados_filtrados = dados[dados['IDADE'] > 10]  # Excluir idades 0
    plt.figure(figsize=(8, 6))
    sns.histplot(dados_filtrados['IDADE'], kde=True)
    plt.title('Distribuição de Idade dos Clientes (Sem Idade 0)')
    plt.xlabel('Idade')
    plt.ylabel('Frequência')
    idade_grafico = 'idade_dist.png'
    plt.savefig(idade_grafico)
    relatorio_html += f"<h2>Distribuição de Idade dos Clientes</h2><img src='{idade_grafico}' alt='Distribuição de Idade'/>"

    # --- 10. Tendência de Vendas ao Longo do Tempo ---
    vendas_mensais = dados.groupby(dados['MÊS_VENDA'].dt.to_period('M'))['QTD_UNIDADE_FARMACOTECNICA'].sum()
    plt.figure(figsize=(8, 6))
    sns.lineplot(x=vendas_mensais.index.astype(str), y=vendas_mensais.values)
    plt.title('Tendência de Vendas por Mês')
    plt.xlabel('Mês')
    plt.ylabel('Quantidade Vendida')
    plt.xticks(rotation=45)
    vendas_grafico = 'vendas_mensais.png'
    plt.savefig(vendas_grafico)
    relatorio_html += f"<h2>Tendência de Vendas ao Longo do Tempo</h2><img src='{vendas_grafico}' alt='Tendência de Vendas'/>"

    # --- 11. Correlação entre Idade e Quantidade Vendida ---
    plt.figure(figsize=(8, 6))
    sns.scatterplot(x=dados_filtrados['IDADE'], y=dados_filtrados['QTD_UNIDADE_FARMACOTECNICA'])
    plt.title('Correlação entre Idade e Quantidade Vendida')
    plt.xlabel('Idade')
    plt.ylabel('Quantidade Vendida')
    correlacao_grafico = 'correlacao_idade_vendas.png'
    plt.savefig(correlacao_grafico)
    relatorio_html += f"<h2>Correlação entre Idade e Quantidade Vendida</h2><img src='{correlacao_grafico}' alt='Correlação entre Idade e Quantidade Vendida'/>"

    # Finalizando o relatório HTML
    relatorio_html += f"<h2>Conclusão</h2><p>Este relatório apresentou uma análise detalhada das vendas, incluindo tendências ao longo do tempo, características dos clientes (como idade e sexo), e distribuição geográfica das vendas. A análise sugere que...</p>"

    # Salvando o relatório em um arquivo HTML
    with open("relatorio_vendas.html", "w") as f:
        f.write(relatorio_html)

    print("Relatório gerado com sucesso em 'relatorio_vendas.html'")

# Carregar o dataset
dados = pd.read_excel('dados_vendas.xlsx')

# Gerar o relatório
gerar_relatorio(dados)

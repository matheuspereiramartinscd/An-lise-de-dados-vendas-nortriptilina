import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Função para remover acentos (caso precise)
import unicodedata
def remover_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

# Carregar o dataset limpo
dados = pd.read_excel('dados_vendas_limpo.xlsx')

# Criando o documento Word
doc = Document()
doc.add_heading('Relatório de Vendas - Nortriptilina', 0)

# Estilo para títulos
style_title = doc.styles['Heading 1']
style_title.font.bold = True
style_title.font.size = Pt(16)

# Estilo para subtítulos
style_subtitle = doc.styles['Heading 2']
style_subtitle.font.bold = True
style_subtitle.font.size = Pt(14)

# --- Métricas e Gráficos ---

# 1. Total de vendas por mês
vendas_mensais = dados.groupby(['ANO_VENDA', 'MÊS_VENDA'])['QTD_UNIDADE_FARMACOTECNICA'].sum().reset_index()

# Adicionando total de vendas por mês ao documento
doc.add_heading('Total de Vendas por Mês/Ano', level=1)
doc.add_paragraph(f"Total de vendas por mês e ano:")
table = doc.add_table(rows=1, cols=len(vendas_mensais.columns))
hdr_cells = table.rows[0].cells
for i, column in enumerate(vendas_mensais.columns):
    hdr_cells[i].text = column

for index, row in vendas_mensais.iterrows():
    row_cells = table.add_row().cells
    for i, value in enumerate(row):
        row_cells[i].text = str(value)

# Estilizando a tabela
for row in table.rows:
    for cell in row.cells:
        # Ajustando o texto e alinhamento da célula
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.font.size = Pt(10)  # Corrigido: ajustando o tamanho da fonte

# Gráfico de vendas por mês
plt.figure(figsize=(10, 6))
sns.lineplot(data=vendas_mensais, x='MÊS_VENDA', y='QTD_UNIDADE_FARMACOTECNICA', hue='ANO_VENDA', marker='o')
plt.title('Tendência de Vendas por Mês/Ano')
plt.xlabel('Mês/Ano')
plt.ylabel('Quantidade Vendida')
plt.xticks(rotation=45)
plt.tight_layout()
grafico_vendas_mes = 'vendas_por_mes.png'
plt.savefig(grafico_vendas_mes)
plt.close()

doc.add_paragraph("Gráfico de Tendência de Vendas por Mês/Ano:")
doc.add_picture(grafico_vendas_mes, width=Inches(5.5))

# 2. Vendas por Gênero
vendas_genero = dados.groupby('GÊNERO')['QTD_UNIDADE_FARMACOTECNICA'].sum().reset_index()

# Adicionando vendas por gênero ao documento
doc.add_heading('Vendas por Gênero', level=1)
doc.add_paragraph(f"Vendas distribuídas por Gênero:")
table = doc.add_table(rows=1, cols=len(vendas_genero.columns))
hdr_cells = table.rows[0].cells
for i, column in enumerate(vendas_genero.columns):
    hdr_cells[i].text = column

for index, row in vendas_genero.iterrows():
    row_cells = table.add_row().cells
    for i, value in enumerate(row):
        row_cells[i].text = str(value)

# Gráfico de vendas por gênero
plt.figure(figsize=(8, 6))
sns.barplot(data=vendas_genero, x='GÊNERO', y='QTD_UNIDADE_FARMACOTECNICA', palette='muted')
plt.title('Vendas por Gênero')
plt.xlabel('Gênero')
plt.ylabel('Quantidade Vendida')
plt.tight_layout()
grafico_vendas_genero = 'vendas_por_genero.png'
plt.savefig(grafico_vendas_genero)
plt.close()

doc.add_paragraph("Gráfico de Vendas por Gênero:")
doc.add_picture(grafico_vendas_genero, width=Inches(5.5))

# 3. Vendas por Estado (UF)
vendas_uf = dados.groupby('UF_VENDA')['QTD_UNIDADE_FARMACOTECNICA'].sum().reset_index()

# Adicionando vendas por estado ao documento
doc.add_heading('Vendas por Estado (UF)', level=1)
doc.add_paragraph(f"Vendas distribuídas por Estado (UF):")
table = doc.add_table(rows=1, cols=len(vendas_uf.columns))
hdr_cells = table.rows[0].cells
for i, column in enumerate(vendas_uf.columns):
    hdr_cells[i].text = column

for index, row in vendas_uf.iterrows():
    row_cells = table.add_row().cells
    for i, value in enumerate(row):
        row_cells[i].text = str(value)

# Gráfico de vendas por estado
plt.figure(figsize=(12, 8))
sns.barplot(data=vendas_uf, x='UF_VENDA', y='QTD_UNIDADE_FARMACOTECNICA', palette='viridis')
plt.title('Vendas por Estado (UF)')
plt.xlabel('Estado')
plt.ylabel('Quantidade Vendida')
plt.xticks(rotation=90)
plt.tight_layout()
grafico_vendas_uf = 'vendas_por_uf.png'
plt.savefig(grafico_vendas_uf)
plt.close()

doc.add_paragraph("Gráfico de Vendas por Estado (UF):")
doc.add_picture(grafico_vendas_uf, width=Inches(5.5))

# 4. Vendas por Município
vendas_municipio = dados.groupby('MUNICIPIO_VENDA')['QTD_UNIDADE_FARMACOTECNICA'].sum().reset_index()

# Adicionando vendas por município ao documento
doc.add_heading('Vendas por Município', level=1)
doc.add_paragraph(f"Vendas distribuídas por Município:")
table = doc.add_table(rows=1, cols=len(vendas_municipio.columns))
hdr_cells = table.rows[0].cells
for i, column in enumerate(vendas_municipio.columns):
    hdr_cells[i].text = column

for index, row in vendas_municipio.iterrows():
    row_cells = table.add_row().cells
    for i, value in enumerate(row):
        row_cells[i].text = str(value)

# Gráfico de vendas por município
plt.figure(figsize=(12, 8))
sns.barplot(data=vendas_municipio.head(10), x='MUNICIPIO_VENDA', y='QTD_UNIDADE_FARMACOTECNICA', palette='coolwarm')
plt.title('Top 10 Vendas por Município')
plt.xlabel('Município')
plt.ylabel('Quantidade Vendida')
plt.xticks(rotation=90)
plt.tight_layout()
grafico_vendas_municipio = 'vendas_por_municipio.png'
plt.savefig(grafico_vendas_municipio)
plt.close()

doc.add_paragraph("Gráfico de Vendas por Município (Top 10):")
doc.add_picture(grafico_vendas_municipio, width=Inches(5.5))

# 5. Estatísticas da Idade
doc.add_heading('Estatísticas de Idade', level=1)
doc.add_paragraph(f"Estatísticas de Idade dos Clientes:")
idade_stats = dados['IDADE'].describe()
doc.add_paragraph(idade_stats.to_string())

# --- SALVANDO O RELATÓRIO ---

# Salvar o documento
relatorio_path = 'relatorio_vendas_nortriptilina.docx'
doc.save(relatorio_path)

print(f"Relatório gerado e salvo em: {relatorio_path}")


# Análise de Dados de Vendas - Cloridrato de Nortriptilina

Este repositório contém scripts Python para análise e limpeza de dados relacionados às vendas do medicamento Cloridrato de Nortriptilina.

## Objetivo

O objetivo deste projeto é processar e analisar um dataset de vendas, fornecendo insights úteis para tomada de decisão. O foco está na limpeza de dados, análise estatística e geração de relatórios.

## Estrutura do Projeto

### 1. `limpeza_dados.py`

- **Descrição**: Script para tratamento e limpeza de dados.
- **Funcionalidades**:
  - Identificação e tratamento de valores ausentes.
  - Exclusão de duplicatas.
  - Padronização de datas e valores monetários.
  - Exclusão de registros inválidos, como idades menores que 18.

### 2. `analise_vendas.py`

- **Descrição**: Script para análise exploratória de dados.
- **Principais Análises**:
  - Cálculo do faturamento total e quantidade de produtos vendidos.
  - Estatísticas de vendas por mês (média, mediana e desvio padrão).
  - Distribuição das vendas por sexo, estado e município.
  - Visualização da distribuição etária dos clientes e tendência de vendas ao longo do tempo.
  - Análise de outliers e correlações entre variáveis.

### 3. `relatorio.py`

- **Descrição**: Script para geração de relatórios em formato Word, com visualizações gráficas geradas a partir do dataset.
- **Principais Funcionalidades**:
  - Geração de tabelas e gráficos para vendas por sexo, estado e município.
  - Visualizações de tendências e distribuições, como correlação entre idade e quantidade vendida.

## Dataset

O dataset utilizado contém as seguintes colunas:

| Coluna                     | Descrição                          |
|----------------------------|------------------------------------|
| `ID`                       | Identificador único da venda       |
| `PRINCIPIO_ATIVO`          | Nome do princípio ativo            |
| `DCB`                      | Código DCB                        |
| `ANO_VENDA`                | Ano da venda                      |
| `MÊS_VENDA`                | Data do mês da venda              |
| `MÊS_VENCIMENTO`           | Data do vencimento                |
| `SEXO`                     | Sexo do cliente                   |
| `IDADE`                    | Idade do cliente                  |
| `UF_VENDA`                 | Estado da venda                   |
| `MUNICIPIO_VENDA`          | Município da venda                |
| `QTD_UNIDADE_FARMACOTECNICA` | Quantidade de unidades vendidas   |
| `TIPO_RECEITUARIO`         | Tipo de receituário               |

## Requisitos

- **Bibliotecas Python**:
  - `pandas`
  - `seaborn`
  - `matplotlib`
  - `docx`

## Como Utilizar

1. Certifique-se de ter o Python instalado na sua máquina (versão 3.7 ou superior).
2. Crie o ambiente virtual: Na raiz do seu projeto, use o seguinte comando para criar o ambiente virtual:                      ```bash   
   python -m venv venv
   ```
3. Ative o ambiente virtual:                                                                                                
   ```bash   
   venv\Scripts\activate
   ```
4. Instale as dependências usando o comando:
5. 
   ```bash
   pip install -r requirements.txt
   ```
6. Execute o script de limpeza de dados:
   ```bash
   python limpeza_dados.py
   ```
7. Execute o script de análise para obter insights:
   ```bash
   python analise_vendas.py
   ```
8. Gere o relatório em Word com:
   ```bash
   python relatorio.py
   ```

## Insights Obtidos

- Identificação de tendências de vendas por mês e ano.
- Análise geográfica detalhada das vendas (por estado e município).
- Entendimento do perfil dos clientes (idade e gênero).
- Identificação de outliers em vendas mensais.



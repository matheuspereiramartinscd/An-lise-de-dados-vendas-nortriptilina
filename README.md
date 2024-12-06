
# Análise de Dados de Vendas - Medicamentos
Este repositório contém scripts Python para análise e limpeza de dados relacionados às vendas de medicamentos
## Objetivo

O objetivo deste projeto é processar e analisar um dataset de vendas, fornecendo insights úteis para tomada de decisão. O foco está na limpeza de dados, análise estatística e geração de relatórios.

https://colab.research.google.com/drive/1Dwdio-B_cC_xovea7SbTMza3C1HSV8Fq?usp=sharing

## Estrutura do Projeto

### 1. `limpeza_dados.py`

- **Descrição**: Script para tratamento e limpeza de dados.
- **Funcionalidades**:
  - Identificação e tratamento de valores ausentes.
  - Exclusão de duplicatas.
  - Padronização de datas e valores monetários.
  - Exclusão de registros inválidos
### 2. `relatorio.py`

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
2. Crie o ambiente virtual: Na raiz do seu projeto, use o seguinte comando para criar o ambiente virtual: 

   ```bash   
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
7. Gere o relatório em Word com:
   ```bash
   python relatorio.py
   ```

## Insights Obtidos

- Identificação de tendências de vendas por mês e ano.
- Análise geográfica detalhada das vendas (por estado e município).
- Identificação de outliers em vendas mensais.

## Dados
![Screenshot_3](https://github.com/user-attachments/assets/8233f5a4-bd6c-44b2-b15a-4bfa0ebdae49)


## Após limpeza
![Screenshot_2](https://github.com/user-attachments/assets/07715a86-b79d-45c1-8dd2-4598d507d5f4)


## Relatório
![Screenshot_4](https://github.com/user-attachments/assets/53eb7cd3-c623-4ef8-90fd-aa4666139279)
![Screenshot_5](https://github.com/user-attachments/assets/827650de-6550-420f-822d-8a210293c124)
![relatorio_vendas_nortriptilina-10](https://github.com/user-attachments/assets/fb350606-e67e-45bf-980e-01b5c57dd1cc)
![relatorio_vendas_nortriptilina-09](https://github.com/user-attachments/assets/dca826d0-3c5e-4822-8b72-54e1a5dc9c83)
![relatorio_vendas_nortriptilina-08](https://github.com/user-attachments/assets/3ecfd0d3-c716-42f9-96f3-e5a897cfa768)
![relatorio_vendas_nortriptilina-07](https://github.com/user-attachments/assets/3c2d7df4-fd71-495a-91cf-63435e6e62be)
![relatorio_vendas_nortriptilina-06](https://github.com/user-attachments/assets/06129051-5c7b-4491-a795-858d28782d8a)
![relatorio_vendas_nortriptilina-05](https://github.com/user-attachments/assets/78986b48-ae61-4381-bf85-6b01d0c665b8)

## Dashboard criado no PowerBI
![Screenshot_1](https://github.com/user-attachments/assets/04971763-d331-4007-8cc1-c1bd27db1c16)






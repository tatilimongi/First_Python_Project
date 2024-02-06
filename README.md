# Automação de Planilhas com Python

Este repositório contém um estudo de caso sobre automação de planilhas utilizando Python para a XYZ Corporation International, uma revendedora de automóveis de luxo sediada em São Paulo. O objetivo é apresentar os resultados da equipe comercial para o novo CEO da empresa, fornecendo informações sobre as vendas de carros por fabricante ao longo dos anos.

## Fonte de Dados

Os dados são provenientes de um arquivo Excel que contém informações coletadas do sistema de vendas e CRM da empresa. As colunas presentes no arquivo são:

| Coluna          | Descrição                                           |
|-----------------|-----------------------------------------------------|
| DataNotaFiscal | Data de emissão da nota fiscal                      |
| Fabricante      | Fabricante do veículo                               |
| Estado          | Estado onde foi realizada a venda                    |
| PrecoVenda      | Preço de venda do veículo                           |
| PrecoCusto      | Preço de custo do veículo para a empresa            |
| TotalDesconto   | Total de desconto fornecido sobre o preço de venda  |
| CustoEntrega    | Custo de entrega do veículo ao proprietário         |
| CustoMaoDeobra  | Custo de mão de obra (secretária, mecânico, etc...) |
| NomeCliente     | Nome do cliente que comprou o veículo               |
| Modelo          | Modelo do veículo                                   |
| Ano             | Ano de fabricação do veículo                        |

## Objetivo da Automação

A automação deve seguir os seguintes passos:

1. Gerar uma tabela pivô em Python contendo as colunas: Fabricante, PrecoVenda e Ano, onde a coluna Ano será o índice para a nova tabela.
2. Importar a nova planilha gerada e criar um gráfico de barras para apresentar o total de vendas por fabricante ao longo dos anos.
3. Adicionar fórmulas para totalizar as vendas por cada fabricante.
4. Enviar a planilha automatizada por e-mail para o CEO da empresa.

## Instruções de Uso

1. Certifique-se de ter o Python instalado em sua máquina.
2. Clone este repositório em sua máquina local.
3. Execute o script Python `automacao_planilhas.py` para gerar a nova planilha e o gráfico.
4. Verifique o arquivo `resultado_automacao.xlsx` para os resultados gerados.
5. Configure o envio automático por e-mail conforme necessário.

## Bibliotecas Utilizadas

- Pandas: Para manipulação de dados e geração da tabela pivô.
- Matplotlib: Para criação do gráfico de barras.

## Autor

Este estudo de caso foi realizado por Tatiana Limongi Chaves através da orientação do OneBitCode.

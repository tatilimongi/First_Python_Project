import pandas as pd

# Importar dados
data = pd.read_excel('data/VendaCarros.xlsx')
# print(type(data))

# Selecionar colunas especificas do dataframe
df = data[['Fabricante', 'ValorVenda', 'Ano']]
# print(df)

# Criar tabela pivo
pivot_table = df.pivot_table(
    index = 'Ano',
    columns = 'Fabricante',
    values = 'ValorVenda',
    aggfunc = 'sum'
)

print(pivot_table)

# Exportar tabela pivo em arquivo excel
pivot_table.to_excel('data/pivot_table.xlsx', 'Relatorio')
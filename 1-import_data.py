import pandas as pd

# Importar dados
data = pd.read_excel('data/VendaCarros.xlsx')
print(data)

# Listar os primeiros registros
print(data.head())

# Listar ultimos registros
print(data.tail())

# Contagem de valores por Fabricante
print(data['Fabricante'].value_counts())

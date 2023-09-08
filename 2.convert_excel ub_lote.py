import pandas as pd
pd.options.display.max_columns = None
pd.options.display.max_rows = None 
# Leer el archivo de Excel
df = pd.read_excel(r'C:\xampp\htdocs\py\comp\QueryResults298430.xlsx')

# Seleccionar las columnas que desea extraer
df = df[['IMLITM', 'LIMCU', 'LILOCN', 'LILOTN']]

# Eliminar las primeras 10 filas
#df = df.iloc[10:]

# Renombrar las columnas
df.columns = ['item', 'bodega', 'ubicacion', 'lote']

# Convertir la columna a tipo cadena
#df['IMLITM'] = df['item'].astype(str)
# Convertir la columna a tipo cadena

# Eliminar los espacios en blanco de cada fila

df = df.dropna(how='all')
df = df.dropna(subset=['ubicacion', 'lote'], how='all')


df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
# Guardar el resultado en un nuevo archivo de Excel
df.to_excel(r'C:\xampp\htdocs\py\comp\lleno.xlsx', index=False)

# Imprimir el resultado
print(df)
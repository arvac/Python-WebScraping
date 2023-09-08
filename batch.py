import pandas as pd
import numpy as np

pd.options.display.max_columns = None
pd.options.display.max_rows = None 
pd.set_option('display.max_colwidth', -1)
# Leer el archivo de Excel
df = pd.read_excel(r'C:\xampp\htdocs\py\comp\tabula.xlsx', header=None)

# Agregar las cabeceras deseadas
df.columns = ["Columna1", "Columna2", "Columna3", "Columna4", "Columna5", "Columna6", "Columna7", "Columna8", "Columna9"]

# Eliminar los espacios en blanco de cada fila
df = df.dropna(how='all')
df = df.dropna(subset=['Columna7'], how='all')
df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

# Aplicar la fórmula en las columnas 5 y 6
df["Columna5"] = np.where(df["Columna2"] == df["Columna2"].shift(), "", df.groupby("Columna2")["Columna9"].transform("sum"))
df["Columna6"] = np.where(df["Columna2"] == df["Columna2"].shift(), "", df.groupby("Columna2")["Columna8"].transform("sum"))

df["Columna6"] = pd.to_numeric(df["Columna6"], errors="coerce")
df["Columna5"] = pd.to_numeric(df["Columna5"], errors="coerce")
df["Columna5"] = df["Columna5"].round(2)
df["Columna6"] = df["Columna6"].round(2)
df["Columna4"] = df["Columna5"] - df["Columna6"]

# Guardar el resultado en un nuevo archivo de Excel 
df.to_excel(r'C:\xampp\htdocs\py\comp\tabula1.xlsx', index=False)

# Imprimir el resultado =SI(B2=B1," ",SUMAR.SI(B:B,B2,I:I)) =SI(B2=B1," ",SUMAR.SI(B:B,B2,H:H))
print(df)
'''
# Convierte los valores no numéricos en NaN
df["Columna6"] = pd.to_numeric(df["Columna6"], errors="coerce")

# Elimina las filas con valores NaN en la columna "Columna6"
df = df.dropna(subset=["Columna6"])

# Convierte los valores no numéricos en NaN
df["Columna5"] = pd.to_numeric(df["Columna5"], errors="coerce")

# Elimina las filas con valores NaN en la columna "Columna6"
df = df.dropna(subset=["Columna5"])

# Convierte los valores no numéricos en NaN
df["Columna4"] = pd.to_numeric(df["Columna4"], errors="coerce")

# Elimina las filas con valores NaN en la columna "Columna6"
df = df.dropna(subset=["Columna6"])


# Convertir la columna a tipo cadena
df["Columna4"] = df.groupby("Columna5")["Columna6"].transform(lambda x: x - x.iloc[0])

'''
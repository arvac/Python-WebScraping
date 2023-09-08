import pandas as pd
from pandasql import sqldf

pd.options.display.max_columns = None
pd.options.display.max_rows = None 
pd.set_option('display.max_colwidth', -1)
pysqldf = lambda q: sqldf(q, globals())
# Cargar los archivos de Excel en DataFrames
df2 = pd.read_excel(r'C:\xampp\htdocs\py\comp\comprobantes.xlsx')
df1 = pd.read_excel(r'C:\xampp\htdocs\py\comp\mensajes.xlsx')



# Crear una conexión a una base de datos en memoria
# Realiza una consulta SQL para combinar los datos de ambos DataFrames
resultados1 = pysqldf('SELECT df2.cia, df2.doc , df2.numero, df2.emision   FROM df2 ')
#resultados = pysqldf('SELECT df1.cia, df1.doc, df1.numero,  df1.informacion from df1 JOIN df2 ON df1.cia = df2.cia AND df1.doc = df2.doc AND df2.numero = df2.numero where df1.numero = "176227" ;')
resultados = pysqldf('SELECT df1.cia, df1.doc, df1.numero,  df1.informacion from df1 JOIN df2 ON df1.cia = df2.cia AND df1.doc = df2.doc AND df2.numero = df2.numero  ;')
resultados.to_excel(r'C:\xampp\htdocs\py\comp\query.xlsx', index=False)
#resultados = pysqldf('SELECT * FROM df2 WHERE cia="143"')
#resultados = pysqldf('SELECT * FROM df2')
# Muestra los resultado
#print(resultados )
print(resultados )


'''
SELECT cia, doc, numero, NULL AS mensaje
FROM tabla1
UNION ALL
SELECT cia, doc, numero, mensaje
FROM tabla2;

SELECT tabla1.cia, tabla1.doc, tabla1.numero, tabla2.mensaje
FROM tabla1
JOIN tabla2
ON tabla1.cia = tabla2.cia AND tabla1.doc = tabla2.doc AND tabla1.numero = tabla2.numero;

'''
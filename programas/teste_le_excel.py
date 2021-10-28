import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

df = pd.read_excel('\\\\srv-arquivos07\dirgerti\SEDT_AUTORE\TN_AUTORE\SUSTENTAÇÃO\Incidentes\DE_PARA_TRIAGEM.xlsx', shee_tname='Planilha1')

print(df)

df.sort_values(by=['PALAVRAS'] , inplace=True, ascending=False)
#print("======Column headings========\n")
#print(df.columns)
#print('Print palavras')
palavra = 'EQUIPAMENTO'
palavras = df['PALAVRAS']
df.tail()


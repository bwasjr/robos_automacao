import xml.etree.ElementTree as ET
import pandas as pd
import numpy as np


df = pd.read_excel('//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx') #arquivo convertido em xlsx
nulo = [np.nan]
df = df[~df['Explicação'].isin(nulo)]

qtd_incidentes = df['Explicação'].shape[0]
lista_incidentes = []
df['resolucao']=''

print('Tipo de qt_incidentes {}, Tipo lista incidentes {} e len de qtd_incidentes {}'.format(type(qtd_incidentes), type(lista_incidentes), qtd_incidentes))


for x in df['Explicação'].index:
    if('<resolucao>11' in df['Explicação'][x].replace(' ', '')):
            lista_incidentes.append(df['ID do Incidente'][x])
            df['resolucao'][x]=11
            #print(template)                

print(len(lista_incidentes))

df = df[df['ID do Incidente'].isin(lista_incidentes)]#o data frame passa a ter somente os INs com template preenchido

df.to_excel('//SRV-ARQUIVOS07/DirGerTI/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Washington/crivo.xlsx', 'Planilha1',index=False)


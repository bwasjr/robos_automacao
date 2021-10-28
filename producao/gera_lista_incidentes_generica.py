import numpy as np
import time
from datetime import date
from datetime import timedelta
import os
import glob
import pandas as pd
import csv
from xlsxwriter.workbook import Workbook


df = pd.read_excel('C:/Users/g571602/Downloads/base dashboard incidentes.xlsx') #arquivo convertido em xlsx

df['Descrição Resumida'] = df['Descrição Resumida'].str.upper()
df['Descrição'] = df['Descrição'].str.upper()

df = df[~df['Explicação'].isin([np.nan])]

palavras = ['DGTZ','CHASSI','NASCIMENTO','CRIVO','CONDUTOR']

lista_matches = []
df['resolucao']=''

for x in df['Explicação'].index:
    string = df['Explicação'][x].replace(' ', '').replace(' ', '').replace('\n', '').replace('\r', '').replace('\t', '')
    if('<resolucao>11' in string):
            lista_matches.append(df['ID do Incidente'][x])
            df['resolucao'][x]=11
            #print(template)

df = df[df['ID do Incidente'].isin(lista_matches)]

lista_matches.clear()

for x in df['ID do Incidente'].index:
    for palavra in palavras:
        if((palavra in df['Descrição Resumida'][x]) ):
            lista_matches.append(df['ID do Incidente'][x])#guarda a linha
            df['resolucao'][x]=11
            
        else:
            if(type(df['Descrição'][x])!=str):
                print('A descrição do incidente: '+ str(df['ID do Incidente'][x]) + ' está nula' )
            elif(palavra in df['Descrição'][x]):
                lista_matches.append(df['ID do Incidente'][x])#guarda a linha
                df['resolucao'][x]=11
                
df = df[df['ID do Incidente'].isin(lista_matches)]

df.to_excel('//SRV-ARQUIVOS07/DirGerTI/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Washington/crivo.xlsx', 'Planilha1',index=False)

print('arquivo gerado')
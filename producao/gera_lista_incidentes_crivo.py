from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import numpy as np
import time
from datetime import date
from datetime import timedelta
import os
import glob
import pandas as pd
import csv
from xlsxwriter.workbook import Workbook


df = pd.read_excel('//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx') #arquivo convertido em xlsx

df['Descrição Resumida'] = df['Descrição Resumida'].str.upper()
df['Descrição'] = df['Descrição'].str.upper()

#nulo = [np.nan]
#df = df[~df['Descrição'].isin(nulo)]

palavras = ['PLACA','CHASSI','NASCIMENTO','CRIVO','CONDUTOR']

lista_matches = []
qtd_incidentes = df['ID do Incidente'].shape[0]

for x in range(qtd_incidentes):
    for palavra in palavras:
        if(palavra in df['Descrição Resumida'][x]):
            lista_matches.append(df['ID do Incidente'][x])#guarda a linha
        else:
            if(type(df['Descrição'][x])!=str):
                print('A descrição do incidente: '+ str(df['ID do Incidente'][x]) + ' está nula' )
            elif(palavra in df['Descrição'][x]):
                lista_matches.append(df['ID do Incidente'][x])#guarda a linha

print('Tamanho do df original {}'.format(df.shape[0]))
print('Quantidade de matches {}'.format(len(lista_matches)))
print('Quantidade de matches sem duplicatas {}'.format(len(set(lista_matches))))

df = df[df['ID do Incidente'].isin(lista_matches)]

print('Data frame após a remoção dos que não são match {}'.format(df.shape[0]))

df.to_excel('//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/crivo.xlsx', index=False) #arquivo convertido em xlsx
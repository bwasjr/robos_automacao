import numpy as np
import time
from datetime import date
from datetime import timedelta
import os
import glob
import pandas as pd
import csv
from xlsxwriter.workbook import Workbook

df_horas = pd.read_csv('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Washington\\horas_trabalhadas_consultorias\\horas.csv', sep=';', header='infer')
df_horas['tempo_segundos'] = '' #cria a coluna de segundos
df_horas['RS'] = '' #cria a coluna que receberá o nome da RS

#cria coluna de segundos
for indice in range (len(df_horas['Tempo Atividade'])):
    tamanho = len(df_horas['Tempo Atividade'][indice])
    tamanho_substr = tamanho - 8
    tempo = df_horas['Tempo Atividade'][indice]
    horas_segundos = int(tempo[tamanho_substr:tamanho_substr+2])*3600
    minutos_segundos = int(tempo[tamanho_substr+3:tamanho_substr+5])*60
    segundos = int(tempo[-2:])
    total_segundos = horas_segundos + minutos_segundos + segundos
    
    #print('Dias_segundos: {} Horas_segundos: {} Minutos_segundos {} Segundos: {}'.format(dias_segundos, horas_segundos, minutos_segundos, segundos))
    if (tamanho_substr > 0):
            dias_segundos = int(tempo[0:tamanho_substr])*24*3600
            total_segundos+=dias_segundos
    df_horas['tempo_segundos'][indice] = total_segundos

#preenche a coluna com os nomes das RSs
for indice in range(len(df_horas['Número'])):
    artefato = df_horas['Número'][indice]
    if ('-' in artefato):
        artefato = artefato[:artefato.find('-')]
        df_horas['RS'][indice] = artefato

#converte os campos de data para o excel
df_horas["Data Inicio Servico"] = pd.to_datetime(df_horas["Data Inicio Servico"], format='%d/%m/%Y %H:%M:%S')#converte os data types para datetime
df_horas["Data Fim Servico"] = pd.to_datetime(df_horas["Data Fim Servico"], format='%d/%m/%Y %H:%M:%S')#converte os data types para datetime

#cria o arquivo excel
df_horas.to_excel('//SRV-ARQUIVOS07/DirGerTI/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Washington/horas_trabalhadas_consultorias/horas_trabalhadas.xlsx', 'Planilha1',index=False)

print('Arquivo gerado')
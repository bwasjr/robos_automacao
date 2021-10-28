import numpy as np
import time
import datetime
import os
import glob
import pandas as pd
import csv
from xlsxwriter.workbook import Workbook
import interacoes_selenium as IS
from sqlalchemy import create_engine
import pymysql
from dateutil.relativedelta import *

now = datetime.datetime.now()
hora = now.strftime("%Y-%m-%d %H:%M:%S")
ano_mes_agora = str(now.strftime("%Y%m"))

#declara os arquivos
arq_extracao = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx'
arq_historico_incidentes = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/historico_incidentes.xlsx'
arq_incidentes_removidos = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/incidentes_removidos.xlsx'

#cria os dataframes dos arquivos
df_extracao = pd.read_excel(arq_extracao)#dataframe do arquivo de extração do SIGS
df_historico_incidentes = pd.read_excel(arq_historico_incidentes)#dataframe com o histórico dos estoques

print('len historico de incidentes ' + str(len(df_historico_incidentes)))

#aplica filtros
df_incidentes_removidos = df_historico_incidentes[df_historico_incidentes['ID_INCIDENTE'].isin(df_extracao['ID do Incidente'])==False]
print('len incidentes removidos ' + str(len(df_incidentes_removidos)))

idx = df_incidentes_removidos.groupby(['ID_INCIDENTE'])['DATA_EXTRACAO'].transform(max) == df_incidentes_removidos['DATA_EXTRACAO']
df_incidentes_removidos = df_incidentes_removidos[idx]
print('len incidentes removidos ' + str(len(df_incidentes_removidos)))

df_incidentes_removidos = df_incidentes_removidos[df_incidentes_removidos['STATUS'].isin(['ENCERRADO'])==False]
print('len incidentes removidos ' + str(len(df_incidentes_removidos)))

#remove as colunas desnecessárias
df_incidentes_removidos = df_incidentes_removidos.drop(['DATA_RESOLUCAO'], axis=1)

df_incidentes_removidos = df_incidentes_removidos.rename(columns={"STATUS":"ULTIMO_STATUS_BARE", "GRUPO_SIGS":"ULTIMO_GRUPO_BARE", "DATA_EXTRACAO":"DATA_ULTIMA_EXTRACAO"})#renomeia as colunas
df_incidentes_removidos['DATA_VERIFICACAO_SAIDA'] = hora #cria coluna
df_incidentes_removidos['DATA_VERIFICACAO_SAIDA'] = pd.to_datetime(df_incidentes_removidos['DATA_VERIFICACAO_SAIDA'], format='%Y-%m-%d %H:%M:%S')#converte os data types para datetime

df_incidentes_removidos.to_excel(arq_incidentes_removidos, 'Planilha1',index=False)

print(df_incidentes_removidos.tail(5))

print('Arquivo de incidentes removidos da BARE gerado')

import os
import glob
import pandas as pd
import numpy as np

os.chdir("//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/")


#lista_arquivos = ["south_america_2000_2010.csv", "north_america_2000_2010.csv"]
lista_arquivos = ["export.csv", "export (1).csv"]

#combine all files in the list
#combined_csv = pd.concat([pd.read_csv(f, encoding="ansi", error_bad_lines=False, delimiter=',') for f in lista_arquivos ])
combined_csv = pd.concat([pd.read_csv(f, encoding="ansi", delimiter=';', quotechar='"', engine='python') for f in lista_arquivos ], join='outer', ignore_index=False, sort=False)

#export to csv
combined_csv.to_csv( "combined_csv.csv", index=False, encoding="ansi")
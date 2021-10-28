import xml.etree.ElementTree as ET
import pandas as pd
import numpy as np


df = pd.read_excel('//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx') #arquivo convertido em xlsx
nulo = [np.nan]
df = df[~df['Explicação'].isin(nulo)]

qtd_incidentes = df['Explicação'].shape[0]
lista_incidentes = []
df['resolucao']=''

string = "<?xml version='1.0' encoding='UTF-8'?> <!--Template de Encerramento de Incidentes - Sustentacao BARE v8--> <template>     <campo>        <resolucao>11</resolucao>        <artefato_associado> </artefato_associado>        <processo_negocio>Acesso aplicação</processo_negocio>        <abrangencia>1</abrangencia>        <nome_projeto> </nome_projeto>        <aplicacao_pgm>https://wwws.bradescoseguros.com.br/GSRE-GestaoIndenizacao/sitePrestador.do</aplicacao_pgm>        <comentarios>Conversamos há pouco com o Sr. Pedro Pires, pelo TEAMS. Trata-se de um problema de perfil de acesso. A prestadora DELPHOS está entrando na operação, a exemplo da prestadora UON. Orientamos, por exemplo, que seja verificado o perfil da UON e criado semelhante à prestadora DELPHOS. Tomamos conhecimento que o assunto já está em andamento, no âmbito técnico com os responsáveis de infra.</comentarios>     </campo> </template>"
count = 0

for x in df['Explicação'].index:
    if('<?xml version' in df['Explicação'][x]):
            string = df['Explicação'][x].replace('\n','')
            parser = ET.XMLParser(encoding="utf-8")
            root = ET.fromstring(string, parser=parser)
            count +=1
            print(count)
            resolucao = root.find('.//resolucao').text
            df['resolucao'][x]=resolucao
            lista_incidentes.append(df['ID do Incidente'][x])

df = df[df['ID do Incidente'].isin(lista_incidentes)]#o data frame passa a ter somente os INs com template preenchido
df.to_excel('//SRV-ARQUIVOS07/DirGerTI/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Washington/crivo.xlsx', 'Planilha1',index=False)

print('fim da execução')
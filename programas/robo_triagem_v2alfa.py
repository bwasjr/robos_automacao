import robo
import importlib
importlib.reload(robo)
from robo import *
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

#robo.main()#extrai os incidentes do grupo sustentacao bare

def loga_pag_pesquisa():#loga e acessa a página de pesquisa
    driver = robo.instancia_driver()
    driver.get('http://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')#acessa a pagina do SIGS
    robo.login(driver)#executa o login
    robo.painel_esquerda(driver)
    return driver

#driver = loga_pag_pesquisa()#loga no SIGS e navega até a área de pesquisar incidentes
df_dpara = pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\DE_PARA_TRIAGEM.xlsx', shee_tname='Planilha1')
df_incidentes = pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\Incidentes\\robo_sigs\\extracao_robo_sust_bare.xlsx', shee_tname='Planilha1')


for descricao in df_incidentes['Descrição Resumida']:
    for palavra in df_dpara['PALAVRAS']:
        if (palavra in descricao):
            palavra_index = df_dpara[df_dpara['PALAVRAS']==palavra].index.values #pega o indice da palavra no arquivo de DE_PARA
            descricao_index = df_incidentes[df_incidentes['Descrição Resumida']==descricao].index.values#pega o index da descricao
            grupo = df_dpara['GRUPO_DESTINO'][palavra_index] #pega o grupo de destino correspondente à palavra
            print('Grupo:' + grupo)
            df_incidentes['GRUPO_DESTINO'][descricao_index] = grupo#escreve o grupo de destino no arquivo de incidentes
            

#abre o arquivo e começa a ler a coluna dos incidentes
    #le o incidente da planilha
    #digita o numero do incidente na planilha e pesquisa
    #le a descrição do incidente
    #procura as palavras chave da tabela de deXpara na descrição do incidente
    #se palavra encontrada então
        #redireciona para o grupo adequado
        #retorna para a pesquisa de incidentes
    #senão, retorna para a pesquisa de incidentes

print("=======================\nFim da triagem\n=======================")
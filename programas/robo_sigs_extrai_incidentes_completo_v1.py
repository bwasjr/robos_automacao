from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import numpy as np
import time
import os
import glob
import pandas as pd
import csv
from xlsxwriter.workbook import Workbook 

def instancia_driver():    #INICIALIZACAO
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument("--test-type")
    options.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": "\\\\srv-arquivos07\dirgerti\SEDT_AUTORE\TN_AUTORE\SUSTENTAÇÃO\Incidentes\Dashboard Incidentes",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    options.add_experimental_option('prefs', prefs)

    driver = webdriver.Chrome('C:\\Users\\g571602\\Documents\\Python\\selenium\\chromedriver.exe', options=options)
    return driver

    #FUNCOES
def clica_id(driver, id):
    element = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.ID, id)))
    element.click()

def clica_classe(driver, classe, posicao):
    lista = driver.find_elements_by_class_name(classe)
    lista[posicao].click() #clica no elemento desejado

def clica_objeto_lista(lista, posicao_objeto):
    lista[posicao_objeto].click()

def insere_texto(elemento, texto):
    elemento.send_keys(texto)

def troca_frame(driver, frame):
    driver.switch_to.frame(frame)

def retorna_objetos(driver, selecao, nome):
    if (selecao == 'id'):
        return driver.find_element_by_id(nome)
    if (selecao == 'class'):
        return driver.find_elements_by_class_name(nome)
    if (selecao == 'name'):
        return driver.find_element_by_name(nome)
    if ((selecao == 'tag')):
        return driver.find_elements_by_tag_name(nome)

def login(driver):
    insere_texto(retorna_objetos(driver, 'id', 'LoginUsername'), 'g571602')
    insere_texto(retorna_objetos(driver, 'id', 'LoginPassword'), 'was7842')
    clica_id(driver, 'loginBtn')

def painel_esquerda(driver):#selecoes do painel da equerda
    clica_id(driver, 'o') #clica na janela automática
    clica_classe(driver, 'x-panel-header', 4)#expande o gerenciar incidentes
    time.sleep(1)
    clica_id(driver, 'ROOT/Gerenciamento de Incidentes/Pesquisar Incidentes') #clica em pesquisar incidentes

def pesquisa_incidentes(driver, grupo, aberto_apos):#iframe do formulario de pesquisa
    lista_objetos = retorna_objetos(driver, 'tag', 'iframe')#obtem a lista dos iframes
    troca_frame(driver, lista_objetos[-2])#seleciona o iframe do formulario de pesquisa do incidente
    time.sleep(1)
    retorna_objetos(driver, 'id', 'X17LabelSpan').click() #seleciona o "ambos"
    retorna_objetos(driver, 'id', 'X114').clear() #apaga o grupo
    retorna_objetos(driver, 'id', 'X31').clear() #apaga o aberto apos
    insere_texto(retorna_objetos(driver, 'id', 'X114'), grupo)#preenche o grupo 'DS - BS - SUSTENTACAO-BARE*'
    insere_texto(retorna_objetos(driver, 'id', 'X31'), aberto_apos)#preenche o aberto após '31/12/2018 23:59:59'
    driver.switch_to.default_content()#retorna ao content default
    driver.find_element(By.XPATH, '//button[text()="Pesquisar"]').click()#botao de pesquisar

def pagina_lista_incidentes(driver):#pagina que contém a lista dos incidentes que serao exportados
    lista_objetos = retorna_objetos(driver, 'tag', 'button')
    clica_objeto_lista(lista_objetos,18) #clica no botao 'mais'
    lista_objetos.clear()
    lista_objetos = retorna_objetos(driver, 'class', 'x-menu-item-text') #retorna os botoes apos clicar em mais
    clica_objeto_lista(lista_objetos,2)# clica em exportar para arquivo de texto
    
def pagina_exportacao(driver):#frame da pagina de exportacao
    lista_objetos = retorna_objetos(driver, 'tag', 'iframe')#obtem a lista de iframes
    troca_frame(driver, lista_objetos[1])#clica no frame do radio buttons
    lista_objetos.clear()
    #clica_id(driver, 'X8Label')#clica no CSV separado por virgulas
    clica_id(driver, 'X10Label')#clica no tabulação
    clica_id(driver, 'X21')#clica no botao verde de ok
    time.sleep(10)

def pagina_exportacao_second_run(driver):#frame da pagina de exportacao
    lista_objetos = retorna_objetos(driver, 'tag', 'iframe')#obtem a lista de iframes
    troca_frame(driver, lista_objetos[1])#clica no frame do radio buttons
    lista_objetos.clear()
    clica_id(driver, 'X2Edit')#clica para remover o cabeçalho das colunas
    #clica_id(driver, 'X8Label')#clica no CSV separado por virgulas
    clica_id(driver, 'X10Label')#clica no tabulação
    clica_id(driver, 'X21')#clica no botao verde de ok
    time.sleep(10)
    

def second_run(driver):
    driver.switch_to.default_content()#retorna ao content default
    driver.find_element(By.XPATH, '//button[text()="Voltar"]').click()#botao de pesquisar
    pesquisa_incidentes(driver, 'DS - BS - SUSTENTACAO-BARE*', '')
    pagina_lista_incidentes(driver)
    pagina_exportacao_second_run(driver)
    return driver

def main_run(grupo, aberto_apos):
    driver = instancia_driver()
    driver.get('http://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')#acessa a pagina do SIGS
    login(driver)#executa o login
    painel_esquerda(driver)
    pesquisa_incidentes(driver, grupo, aberto_apos)
    pagina_lista_incidentes(driver)
    pagina_exportacao(driver)
    return driver

#main do robo
driver = main_run('DS - BS - SEDT-AUTO/RE*', '')
driver = second_run(driver) #chama a execução da pesquisa do segundo grupo
time.sleep(10)
driver.quit()

#executa o merge dos arquivos baixados
filenames = ["//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/export.txt", "//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/export (1).txt"] #arquivos baixados pelo robo
arquivo_merge = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/arq.txt' #arquivo resultante do merge
arquivo_final = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base_incidentes_completa.xlsx' #arquivo convertido em xlsx

with open(arquivo_merge, 'w') as outfile:#cria o arquivo de saída
    for fname in filenames:
        with open(fname) as infile:
            for line in infile:
                outfile.write(line)#escreve cada linha no arquivo de destino

df = pd.read_csv(arquivo_merge, encoding='ansi', sep='\t') #o arquivo tem o encoding ansi, então é necessário marcar isso 
                                                           #juntamente com o delimitador sep='\t' que significa por tab
df.to_excel(arquivo_final, 'Planilha1',index=False)#gera o arquivo de destino sem a coluna de indice que é gerada automaticamente pelo dataframe do pandas

for arquivo in filenames:#remove os arquivos dos dois baixados do sigs
    if os.path.exists(arquivo):
        os.remove(arquivo)

if os.path.exists(arquivo_merge):#remove o arquivo concatenado
    os.remove(arquivo_merge)
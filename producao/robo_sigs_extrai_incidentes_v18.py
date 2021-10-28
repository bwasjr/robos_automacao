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
#import interacoes_selenium as IS

def instancia_driver():#INICIALIZACAO do chromedriver
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument("--test-type")
    options.add_argument("--start-maximized")
    prefs = {'download.prompt_for_download': False, 'download.default_directory': '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\downloads', 'download.directory_upgrade': True}
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome('C:\\Users\\g571602\\Documents\\Python\\robo_bare\\chromedriver_v86.exe', options=options)
    return driver

def clica_id(driver, id):
    element = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.ID, id)))
    element.click()

def clica_id_time(driver, id, timeout):
    element = WebDriverWait(driver, timeout).until(
    EC.element_to_be_clickable((By.ID, id)))
    element.click()

def clica_xpath(driver, xpath):
    element = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()

def clica_objeto(driver, objeto):
    element = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.ID, objeto.get_attribute("id"))))
    element.click()

def clica_classe(driver, classe, posicao):
    lista = driver.find_elements_by_class_name(classe)
    lista[posicao].click() #clica no elemento desejado

def clica_objeto_lista(lista, posicao_objeto):
    lista[posicao_objeto].click()

def clica_por_texto(lista_objetos,texto):
    limite = 0 #variavel de controle de timout
    clicado = False #variavel de controle que determina que o objeto foi clicado
    while (limite <21 and clicado == False):#loop para esperar o objeto ficar disponível para clicar
        for botao in lista_objetos:
            if ((botao.text == texto) and (botao.is_enabled and botao.is_displayed)):
                botao.click()
                clicado = True
        time.sleep(1)
        limite +=1

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
    insere_texto(retorna_objetos(driver, 'id', 'LoginPassword'), 'was7812')
    clica_id(driver, 'loginBtn')

def painel_esquerda(driver):#selecoes do painel da equerda
    try:
        clica_id_time(driver, 'o', 5) #clica na janela automática
    except:
        pass
    time.sleep(2)
    abas = retorna_objetos(driver,'class', 'x-panel-header')#obtem a lista de abas para clicar
    clica_por_texto(abas, 'Gerenciamento de Incidentes')#clica na aba de pesquisa de incidentes
    time.sleep(2)
    clica_id_time(driver, 'ROOT/Gerenciamento de Incidentes/Pesquisar Incidentes',3) #clica em pesquisar incidentes

def pesquisa_incidentes(driver, grupo, aberto_apos, indice_frame, tipo_execucao):#iframe do formulario de pesquisa
    time.sleep(3)
    lista_objetos = retorna_objetos(driver, 'tag', 'iframe')#obtem a lista dos iframes
    troca_frame(driver, lista_objetos[indice_frame])#seleciona o iframe do formulario de pesquisa do incidente
    if tipo_execucao == 2: clica_id_time(driver, 'X13Label',10)#O tipo de execução =2 significa que é apenas triagem de INs. Portanto, só os INs abertos deverão ser baixados
    else : clica_id_time(driver, 'X17Label',3)#caso contrário a opção "Ambos" deve ser clicada e incidentes abertos e encerrados serão baixados
    retorna_objetos(driver, 'id', 'X116').clear() #apaga o grupo
    retorna_objetos(driver, 'id', 'X31').clear() #apaga o aberto apos
    insere_texto(retorna_objetos(driver, 'id', 'X116'), grupo)#preenche o grupo
    insere_texto(retorna_objetos(driver, 'id', 'X31'), aberto_apos)#preenche o aberto após
    driver.switch_to.default_content()#retorna ao content default
    clica_xpath(driver, '//button[text()="Pesquisar"]')#clica em pesquisar

def pagina_lista_incidentes(driver):#pagina que contém a lista dos incidentes que serao exportados
    time.sleep(2)
    lista_objetos = retorna_objetos(driver, 'tag', 'button')
    clica_por_texto(lista_objetos,'Mais') #clica no botao 'mais'
    lista_objetos.clear()
    lista_objetos = retorna_objetos(driver, 'class', 'x-menu-item-text') #retorna os botoes apos clicar em mais
    clica_objeto_lista(lista_objetos,2)# clica em exportar para arquivo de texto

def aguarda_download(arquivo):
    count = 0 #variavel de controle do timeout do download do arquivo
    timeout = 180 #180 segundos
    while (not os.path.exists(arquivo) and count<timeout):
        time.sleep(1)
        count += 1
        if (count == timeout):
            print('O download do arquivo excedeu o timeout de ' + str(timeout) + ' segundos')
    if (count<timeout):
        print('O arquivo foi baixado em ' + str(count) + ' segundos')

def pagina_exportacao(driver, arquivo, cabecalho):#frame da pagina de exportacao
    time.sleep(3)
    lista_objetos = retorna_objetos(driver, 'tag', 'iframe')#obtem a lista de iframes
    troca_frame(driver, lista_objetos[-1])#clica no frame do radio buttons
    lista_objetos.clear()
    lista_objetos = retorna_objetos(driver, 'tag', 'label')#obtem a lista de iframes
    if cabecalho == False:
        clica_id(driver, 'X2Edit')#clica para remover o cabeçalho das colunas    
    clica_id(driver, 'X10Label')#clica no tabulação    
    clica_id(driver, 'X21')#clica no botao verde de ok
    aguarda_download(arquivo)

def pagina_exportacao_second_run(driver, arquivo):#frame da pagina de exportacao
    lista_objetos = retorna_objetos(driver, 'tag', 'iframe')#obtem a lista de iframes
    troca_frame(driver, lista_objetos[1])#clica no frame do radio buttons
    lista_objetos.clear()
    clica_id(driver, 'X2Edit')#clica para remover o cabeçalho das colunas
    clica_id(driver, 'X10Label')#clica no tabulação
    clica_id(driver, 'X21')#clica no botao verde de ok
    aguarda_download(arquivo)

def second_run(driver, grupo, aberto_apos, arquivo, cabecalho, tipo_execucao):
    time.sleep(3)
    driver.switch_to.default_content()#retorna ao content default
    time.sleep(1)
    clica_xpath(driver, '//button[text()="Voltar"]')#clica em voltar
    pesquisa_incidentes(driver, grupo, aberto_apos, -3, tipo_execucao)#-3 é o íncide do iframe. Deve ser -3 na segunda execução em diante
    pagina_lista_incidentes(driver)
    pagina_exportacao(driver, arquivo, cabecalho)
   
def baixa_incidentes_grupo(grupo, aberto_apos, arquivo, cabecalho, tipo_execucao):
    driver = instancia_driver()
    driver.get('https://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')#acessa a pagina do SIGS
    login(driver)#executa o login
    painel_esquerda(driver)
    time.sleep(1)#remover depois
    pesquisa_incidentes(driver, grupo, aberto_apos, -2, tipo_execucao)#-2 no índice frame porque é o frame a ser utilizado na hora de pesquisar os incidentes na primeira execução. -3 na segunda execução em diante
    pagina_lista_incidentes(driver)
    pagina_exportacao(driver, arquivo, cabecalho)
    return driver

def logoff(driver):
    time.sleep(1)
    driver.switch_to.default_content()
    clica_id(driver, 'toolbarUserInfoButtonId')
    lista_objetos = retorna_objetos(driver, 'class','icon-user-logout')
    lista_objetos[0].click()
    alerta = driver.switch_to.alert
    alerta.accept()

def main(tipo_execucao):
    #main do robo
    if tipo_execucao == 1:#faz extração de uma lista completa: grupos da sustentação + projeto + SAP
        df_grupos =  pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_EXTRACAO.xlsx', shee_tname='Plan1')
    elif tipo_execucao == 2:
        df_grupos =  pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_EXTRACAO_TRIAGEM.xlsx', shee_tname='Plan1')#
    elif tipo_execucao == 3:#remover esta opção depois que o usuário tiver acesso a todos os grupos. Depois basta utilizar o tipo de execução 1
        df_grupos =  pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_EXTRACAO_TIPIFICACAO.xlsx', shee_tname='Plan1')#
    df_grupos = df_grupos[df_grupos['ATIVO?'] == 'S']#remove os grupos inativos

    #montar uma lista com os arquivos que serão criados
    filenames = ['//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export.txt']
    for indice in range(len(df_grupos['GRUPOS']) - 1):#ignora o nome do arquivo do primeiro grupo, pois o primeiro foi inserido manualmente na lista filenames
        filenames.append('//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export ('+  str(indice+1) + ').txt')

    #declaração dos arquivos que serão manipulados
    arquivo_merge = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/arq.txt' #arquivo resultante do merge
    arquivo_final = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx' #arquivo convertido em xlsx
    arquivo_final_comum = '//srv-arquivos07/dirgerti/Comum/Incidentes SAP/base dashboard incidentes.xlsx' #coloca o mesmo arquivo no comum para que a equipe do SAP veja
    arquivo_extracao_robo = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/extracao_robo.xlsx'

    #inicia a extração dos incidentes
    cabecalho = True #flag para definir se é a primeira execução. caso tenha mais de um arquivo baixado, é necessário passar a flag para o método de download para que desmarque a opção de baixar o arquivo com cabeçalho das colunas
    
    #calculo do dia do ano passado para reduzir a volumetria das extrações
    hoje = date.today()
    um_ano = timedelta(367)
    dia_ano_passado = str((hoje - um_ano).strftime("%d/%m/%y")) + ' 23:59:59'
    
    driver = baixa_incidentes_grupo(df_grupos['GRUPOS'][0] , dia_ano_passado, filenames[0], cabecalho, tipo_execucao)
    cabecalho = False #depois da primeira execução o cabeçalho não é mais necessário
    for index in range(len(filenames)-1):#inicia a execução do segundo arquivo em diante
        second_run(driver, df_grupos['GRUPOS'][index+1] , dia_ano_passado, filenames[index+1], cabecalho, tipo_execucao)
    time.sleep(2)
    logoff(driver)
    driver.quit() #encerra o driver   

    #faz o merge dos N arquivos de incidentes baixados
    with open(arquivo_merge,  'w', encoding='utf-8') as outfile:#cria o arquivo de saída
        for fname in filenames:
            with open(fname, encoding='utf-8') as infile:
                for line in infile:
                    outfile.write(line)#escreve cada linha no arquivo de destino
    df = pd.read_csv(arquivo_merge, encoding='utf-8', sep='\t') #o arquivo tem o encoding ansi, então é necessário marcar isso 
                                                            #juntamente com o delimitador sep='\t' que significa por tab
   
    #conversões de string para datetime
    df["Hora de Abertura"] = pd.to_datetime(df["Hora de Abertura"], format='%d/%m/%Y %H:%M:%S')#converte os data types para datetime
    df["Hora de Atualização"] = pd.to_datetime(df["Hora de Atualização"], format='%d/%m/%Y %H:%M:%S')#converte os data types para datetime
    df["Hora de Resolução"] = pd.to_datetime(df["Hora de Resolução"], format='%d/%m/%Y %H:%M:%S')#converte os data types para datetime
    df["Hora de Fechamento"] = pd.to_datetime(df["Hora de Fechamento"], format='%d/%m/%Y %H:%M:%S')#converte os data types para datetime
    df["Hora de Reabertura"] = pd.to_datetime(df["Hora de Reabertura"], format='%d/%m/%Y %H:%M:%S')#converte os data types para datetime
    df["Hora do Alerta"] = pd.to_datetime(df["Hora do Alerta"], format='%d/%m/%Y %H:%M:%S')#converte os data types para datetime

    if tipo_execucao in [1,3]:
        df.to_excel(arquivo_final, 'Planilha1',index=False)#gera o arquivo de destino sem a coluna de indice que é gerada automaticamente pelo dataframe do pandas
        df.to_excel(arquivo_final_comum, 'Planilha1',index=False)#gera o arquivo de destino sem a coluna de indice que é gerada automaticamente pelo dataframe do pandas
        df.to_excel(arquivo_extracao_robo, 'Planilha1',index=False)#gera o arquivo de entrada pra triagem do robô
    if tipo_execucao ==2:
        df.to_excel(arquivo_extracao_robo, 'Planilha1',index=False)#gera o arquivo de entrada pra triagem do robô

    #remove os arquivos baixados e de merge
    for arquivo in filenames:#remove os arquivos baixados do sigs
        if os.path.exists(arquivo):
            os.remove(arquivo)
    if os.path.exists(arquivo_merge):#remove o arquivo concatenado
        os.remove(arquivo_merge)
    print("fim da execução")

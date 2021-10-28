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
from statistics import mode  

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
    element = WebDriverWait(driver, 20).until(
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
    if ((selecao == 'xpath')):
        return driver.find_elements_by_xpath(nome)

def login(driver):
    insere_texto(retorna_objetos(driver, 'id', 'LoginUsername'), 'g571602')
    insere_texto(retorna_objetos(driver, 'id', 'LoginPassword'), 'was7801')
    clica_id(driver, 'loginBtn')

def painel_esquerda(driver):#selecoes do painel da equerda
    #clica_id(driver, 'o') #clica na janela automática
    clica_classe(driver, 'x-panel-header', 4)#expande o gerenciar incidentes
    clica_id(driver, 'ROOT/Gerenciamento de Incidentes/Pesquisar Incidentes') #clica em pesquisar incidentes
    

def pesquisa_incidentes(driver, grupo, aberto_apos):#iframe do formulario de pesquisa
    time.sleep(3)
    lista_objetos = retorna_objetos(driver, 'tag', 'iframe')#obtem a lista dos iframes
    troca_frame(driver, lista_objetos[-2])#seleciona o iframe do formulario de pesquisa do incidente
    insere_texto(retorna_objetos(driver, 'id', 'X114'), grupo)#preenche o grupo 'DS - BS - SUSTENTACAO-BARE*'
    insere_texto(retorna_objetos(driver, 'id', 'X31'), aberto_apos)#preenche o aberto após '31/12/2018 23:59:59'
    driver.switch_to.default_content()#retorna ao content default
    clica_xpath(driver, '//button[text()="Pesquisar"]')#clica em pesquisar

def pagina_lista_incidentes(driver):#pagina que contém a lista dos incidentes que serao exportados
    time.sleep(3)
    lista_objetos = retorna_objetos(driver, 'tag', 'button')
    clica_por_texto(lista_objetos,'Mais') #clica no botao 'mais'
    lista_objetos.clear()
    lista_objetos = retorna_objetos(driver, 'class', 'x-menu-item-text') #retorna os botoes apos clicar em mais
    clica_objeto_lista(lista_objetos,2)# clica em exportar para arquivo de texto

def aguarda_download(arquivo):
    limite = 0 #variavel de controle do timeout do download do arquivo
    while (not os.path.exists(arquivo) and limite<21):
        time.sleep(1)
        limite += 1
        if (limite == 21):
            print('O download do arquivo excedeu o timeout de 20 segundos')

def pagina_exportacao(driver, arquivo):#frame da pagina de exportacao
    lista_objetos = retorna_objetos(driver, 'tag', 'iframe')#obtem a lista de iframes
    troca_frame(driver, lista_objetos[1])#clica no frame do radio buttons
    lista_objetos.clear()
    clica_id(driver, 'X10Label')#clica no tabulação
    clica_id(driver, 'X21')#clica no botao verde de ok
    aguarda_download(arquivo)

def main_run(grupo, aberto_apos, arquivo):
    driver = instancia_driver()
    driver.get('http://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')#acessa a pagina do SIGS
    login(driver)#executa o login
    painel_esquerda(driver)
    pesquisa_incidentes(driver, grupo, aberto_apos)
    pagina_lista_incidentes(driver)
    pagina_exportacao(driver, arquivo)
    return driver

def logoff(driver):
    driver.switch_to.default_content()
    lista_objetos = retorna_objetos(driver, 'class','logout')
    lista_objetos[0].click()
    alerta = driver.switch_to.alert
    time.sleep(1)
    alerta.accept()
    time.sleep(1)


#main do robo
def extrai_incidentes():
    filenames = ["//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/export.txt"] #arquivos baixados pelo robo
    print("=======================\nInício de execução do robô de extração\n=======================")
    driver = main_run('DS - BS - SUSTENTACAO-BARE', '31/12/2018 23:59:59', filenames[0])
    time.sleep(3)
    logoff(driver)
    driver.quit()

def classifica():
    #declaração dos arquivos que serão manipulados
    filenames = ["//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/export.txt"] #arquivos baixados pelo robo
    arquivo_merge = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/arq.txt' #arquivo resultante do merge
    arquivo_final = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/extracao_robo_sust_bare.xlsx' #arquivo convertido em xlsx

    with open(arquivo_merge, 'w') as outfile:#cria o arquivo de saída
        for fname in filenames:
            with open(fname) as infile:
                for line in infile:
                    outfile.write(line)#escreve cada linha no arquivo de destino

    df = pd.read_csv(arquivo_merge, encoding='ansi', sep='\t') #o arquivo tem o encoding ansi, então é necessário marcar isso juntamente com o delimitador sep='\t' que significa por tab
    df = df[df['Designação principal'] == 'DS - BS - SUSTENTACAO-BARE'] #remove as linhas que não tenham o grupo DS - BS - SUSTENTACAO-BARE
    status = ['DIRECIONADO', 'EM TRATAMENTO']#lista de status aceitos no arquivo
    df = df[df['Status'].isin(status)]#sustitui o data frame somente com as linhas que contém os status direcionado e em tratamento
    
    #Le as tabelas de de_para
    df_desc = pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\DE_PARA_TRIAGEM_DESCRICAO.xlsx', shee_tname='Planilha1')
    df_tipo = pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\DE_PARA_TRIAGEM_TIPO_PRODUTO.xlsx', shee_tname='Planilha1')

    lista_grupos = []
    count = -1 #contador para acompanhar o indice do dataframe

    for tipo in df['Tipo de Produto']:#pega o tipo de produto da lista de incidentes
        count +=1 
        if (tipo in df_tipo['TIPO_PRODUTO'].values):#se o tipo de produto da lista de incidentes existe na tabela de de_para
            tipo_depara_index = df_tipo[df_tipo['TIPO_PRODUTO']==tipo].index.values#recupera o indice do tipo de produto na depara
            grupo = df_tipo['GRUPO'][tipo_depara_index]#recupera o nome do grupo na depara
            lista_grupos.append(grupo.values[0])#adiciona o nome do grupo na lista de grupos de destino
        else:#tenta classificar com a descrição
            descricao = df.iloc[count, 8]#busca o a descricao do incidente. foi passada a linha com o count e a coluna com o 8
            lista_matches = []#lista para armazenar todos os codigos de grupos que ocorreram matches entre palavras e descricoes
            for palavra in df_desc['PALAVRAS']:
                palavra_index = df_desc[df_desc['PALAVRAS']==palavra].index.values #pega o indice da palavra no arquivo de DE_PARA
                if (palavra in descricao):
                    cd_grupo = int(df_desc['CODIGO_GRUPO'][palavra_index]) #pega o código do grupo de destino correspondente à palavra
                    lista_matches.append(cd_grupo)#adiciona codigo do grupo na lista de matches
            primeiro = max(set(lista_matches), key=lista_matches.count, default=-1) #elege o grupo mais votado
            grupo_index = df_desc[df_desc['CODIGO_GRUPO']==primeiro].index.values #pega o indice da palavra no arquivo de DE_PARA
            if (len(grupo_index)>0):#se houve matches
                grupo = df_desc['GRUPO_DESTINO'][grupo_index[0]]#acessa o nome do grupo pela primeira ocorrencia de indice na lista grupo_index
                lista_grupos.append(grupo)
            else:#se não foram encontradas correspondências de palavras
                lista_grupos.append('INDETERMINADO')
    df['GRUPO_DESTINO'] = lista_grupos #cria a coluna dos grupos de destino no dataframe
    df.to_excel(arquivo_final, 'Planilha1',index=False)#gera o arquivo de destino sem a coluna de indice que é gerada automaticamente pelo dataframe do pandas

    for arquivo in filenames:#remove os baixados do sigs
        if os.path.exists(arquivo):
            os.remove(arquivo)

    if os.path.exists(arquivo_merge):#remove o arquivo concatenado
        os.remove(arquivo_merge)
    
extrai_incidentes()
classifica()

print("=======================\nFim da execução\n=======================")
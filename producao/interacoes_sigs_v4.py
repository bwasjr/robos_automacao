from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import numpy as np
import time
from datetime import date
from datetime import datetime
from datetime import timedelta
import os
import glob
import pandas as pd
import csv
from xlsxwriter.workbook import Workbook
import interacoes_selenium as IS
from sqlalchemy import create_engine
import pymysql

#arquivos globais
triagem_arquivo_entrada = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/extracao_robo.xlsx' #arquivo convertido em xlsx
triagem_arquivo_final = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/classificao_triagem.xlsx' #arquivo com os incidentes classificados

def instancia_driver():#INICIALIZACAO do chromedriver
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument("--test-type")
    options.add_argument("--start-maximized")
    prefs = {'download.prompt_for_download': False, 'download.default_directory': '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\downloads', 'download.directory_upgrade': True}
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome('C:\\Users\\g571602\\Documents\\Python\\robo_bare\\chromedriver.exe', options=options)
    return driver

def login(driver):
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'LoginUsername'), 'g571602')
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'LoginPassword'), 'was7802')
    IS.clica_id(driver, 'loginBtn')

def logoff(driver):
    time.sleep(1)
    driver.switch_to.default_content()
    IS.clica_id(driver, 'toolbarUserInfoButtonId')
    lista_objetos = IS.retorna_objetos(driver, 'class','icon-user-logout')
    lista_objetos[0].click()
    alerta = driver.switch_to.alert
    alerta.accept()
    driver.quit() #encerra o driver

def painel_esquerda(driver, menu, submenu, timeout):#selecoes do painel da equerda e interage com o artefato desejado do SIGS
    time.sleep(2)
    try:
        IS.clica_id_time(driver, 'o', timeout) #clica na janela automática
    except:
        pass
    time.sleep(3)
    lista_objetos = IS.retorna_objetos(driver,'class', 'x-panel-header')#obtem a lista de abas para clicar
    if (menu !=''):
        IS.clica_por_texto_time(lista_objetos, menu, timeout)#clica no item de menu desejado de acordo com o texto
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver,'class', 'x-tree-node-el')
    IS.clica_por_texto_time(lista_objetos, submenu, timeout) #clica no submenu desejado

def pesquisa_incidentes(driver, grupo, aberto_apos, indice_frame, tipo_execucao):#iframe do formulario de pesquisa
    time.sleep(3)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')#obtem a lista dos iframes
    IS.troca_frame(driver, lista_objetos[indice_frame])#seleciona o iframe do formulario de pesquisa do incidente
    time.sleep(2)
    if (tipo_execucao != 1): IS.clica_id_time(driver, 'X13Label',10)#Só baixa os incidentes em aberto
    else : IS.clica_id_time(driver, 'X17Label',10)#caso contrário a opção "Ambos" deve ser clicada e incidentes abertos e encerrados serão baixados
    IS.retorna_objetos(driver, 'id', 'X116').clear() #apaga o grupo
    IS.retorna_objetos(driver, 'id', 'X31').clear() #apaga o aberto apos
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X116'), grupo)#preenche o grupo
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X31'), aberto_apos)#preenche o aberto após
    driver.switch_to.default_content()#retorna ao content default
    time.sleep(2)
    IS.clica_xpath(driver, '//button[text()="Pesquisar"]')#clica em pesquisar

def pagina_lista_artefatos_pesquisados(driver):#pagina que contém a lista dos incidentes que serao exportados
    time.sleep(4)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'button')
    time.sleep(1)
    IS.clica_por_texto(lista_objetos,'Mais') #clica no botao 'mais'
    lista_objetos.clear()
    lista_objetos = IS.retorna_objetos(driver, 'class', 'x-menu-item-text') #retorna os botoes apos clicar em mais
    IS.clica_objeto_lista(lista_objetos,-2)# clica em exportar para arquivo de texto

def aguarda_download(arquivo, timeout):
    count = 0 #variavel de controle do timeout do download do arquivo
    while (not os.path.exists(arquivo) and count<timeout):
        time.sleep(1)
        count += 1
        if (count == timeout):
            print('O download do arquivo excedeu o timeout de ' + str(timeout) + ' segundos')
    if (count<timeout):
        print('O arquivo foi baixado em ' + str(count) + ' segundos')

def pagina_exportacao(driver, arquivo, cabecalho, timeout):#frame da pagina de exportacao
    time.sleep(3)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')#obtem a lista de iframes
    IS.troca_frame(driver, lista_objetos[-1])#clica no frame do radio buttons
    lista_objetos.clear()
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'label')#obtem a lista de iframes
    if cabecalho == False:
        IS.clica_id(driver, 'X2Edit')#clica para remover o cabeçalho das colunas    
    IS.clica_id(driver, 'X10Label')#clica no tabulação    
    IS.clica_id(driver, 'X21')#clica no botao verde de ok
    aguarda_download(arquivo, timeout)

def pagina_exportacao_second_run(driver, arquivo, timeout):#frame da pagina de exportacao
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')#obtem a lista de iframes
    IS.troca_frame(driver, lista_objetos[1])#clica no frame do radio buttons
    lista_objetos.clear()
    IS.clica_id(driver, 'X2Edit')#clica para remover o cabeçalho das colunas
    IS.clica_id(driver, 'X10Label')#clica no tabulação
    IS.clica_id(driver, 'X21')#clica no botao verde de ok
    aguarda_download(arquivo, timeout)

def second_run_incidentes(driver, grupo, aberto_apos, arquivo, cabecalho, tipo_execucao, timeout_arquivo):
    time.sleep(3)
    driver.switch_to.default_content()#retorna ao content default
    time.sleep(3)
    IS.clica_xpath(driver, '//button[text()="Voltar"]')#clica em voltar
    pesquisa_incidentes(driver, grupo, aberto_apos, -3, tipo_execucao)#-3 é o íncide do iframe. Deve ser -3 na segunda execução em diante
    pagina_lista_artefatos_pesquisados(driver)
    pagina_exportacao(driver, arquivo, cabecalho, timeout_arquivo)

def baixa_incidentes_grupo(grupo, aberto_apos, arquivo, cabecalho, tipo_execucao):
    driver = instancia_driver()
    driver.get('https://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')#acessa a pagina do SIGS
    login(driver)#executa o login
    painel_esquerda(driver, 'Gerenciamento de Incidentes', 'Pesquisar Incidentes', 4)
    time.sleep(1)#remover depois
    pesquisa_incidentes(driver, grupo, aberto_apos, -2, tipo_execucao)#-2 no índice frame porque é o frame a ser utilizado na hora de pesquisar os incidentes na primeira execução. -3 na segunda execução em diante
    pagina_lista_artefatos_pesquisados(driver)
    pagina_exportacao(driver, arquivo, cabecalho, 200)
    return driver

def main_extrai_incidentes(tipo_execucao):#principal execução responsável por extrair os incidentes
    
    if tipo_execucao == 1:#faz extração de uma lista completa: grupos da sustentação + projeto + SAP
        df_grupos =  pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_EXTRACAO.xlsx', sheet_name='Plan1')
    elif tipo_execucao == 2:#extração para triagem
        df_grupos =  pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_EXTRACAO_TRIAGEM.xlsx', sheet_name='Plan1')#
    elif tipo_execucao == 3:#extração para tipificação dos incidentes
        df_grupos =  pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_EXTRACAO_TIPIFICACAO.xlsx', sheet_name='Plan1')#
    df_grupos = df_grupos[df_grupos['ATIVO?'] == 'S']#remove os grupos inativos
    
    #montar uma lista com os arquivos que serão criados
    filenames = ['//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export.txt']
    for indice in range(len(df_grupos['GRUPOS']) - 1):#ignora o nome do arquivo do primeiro grupo, pois o primeiro foi inserido manualmente na lista filenames
        filenames.append('//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export ('+  str(indice+1) + ').txt')

    #declaração dos arquivos que serão manipulados
    arquivo_merge = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/arq.txt' #arquivo resultante do merge
    arquivo_final = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx' #arquivo convertido em xlsx
    arquivo_extracao_robo = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/extracao_robo.xlsx'

    
    #deleta arquivos do diretório de download antes de iniciar o download para evitar que erros de execução
    deleta_arquivos_temporarios(filenames, arquivo_merge)
    
    #inicia a extração dos incidentes
    cabecalho = True #flag para definir se é a primeira execução. caso tenha mais de um arquivo baixado, é necessário passar a flag para o método de download para que desmarque a opção de baixar o arquivo com cabeçalho das colunas
    
    #calculo do dia do ano passado para reduzir a volumetria das extrações
    hoje = date.today()
    um_ano = timedelta(367)
    dia_ano_passado = str((hoje - um_ano).strftime("%d/%m/%y")) + ' 23:59:59'
    
    
    driver = baixa_incidentes_grupo(df_grupos['GRUPOS'][0] , dia_ano_passado, filenames[0], cabecalho, tipo_execucao)
    cabecalho = False #depois da primeira execução o cabeçalho não é mais necessário
    for index in range(len(filenames)-1):#inicia a execução do segundo arquivo em diante
        second_run_incidentes(driver, df_grupos['GRUPOS'][index+1] , dia_ano_passado, filenames[index+1], cabecalho, tipo_execucao, 200)#espera até 200 segundos pelo arquivo
    time.sleep(2)
    logoff(driver)
    

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

    #remoção de aspas no id do incidente
    df['ID do Incidente'] = df['ID do Incidente'].str.replace('"','')
    
    if tipo_execucao in [1,3]:
        df.to_excel(arquivo_final, 'Planilha1',index=False)#gera o arquivo de destino sem a coluna de indice que é gerada automaticamente pelo dataframe do pandas
        df.to_excel(arquivo_extracao_robo, 'Planilha1',index=False)#gera o arquivo de entrada pra triagem do robô
    if tipo_execucao ==2:
        df.to_excel(arquivo_extracao_robo, 'Planilha1',index=False)#gera o arquivo de entrada pra triagem do robô

    #remove os arquivos baixados e de merge
    deleta_arquivos_temporarios(filenames, arquivo_merge)

    print('Fim da extração de incidentes')

def triagem_classifica():
    print("========================================Início da classificação========================================")
    df = pd.read_excel(triagem_arquivo_entrada) #o arquivo tem o encoding ansi, então é necessário marcar isso juntamente com o delimitador sep='\t' que significa por tab
    df = df[df['Designação principal'] == 'DS - BS - SUSTENTACAO-BARE'] #remove as linhas que não tenham o grupo DS - BS - SUSTENTACAO-BARE
    df = df[df['Brd Tipo Ambiente'] == 'PRODUCAO'] #faz triagem somente dos incidentes de produção. Os de outros ambientes não são tratados pela sustentação
    status = ['DIRECIONADO']#lista de status aceitos no arquivo
    df = df[df['Status'].isin(status)]#sustitui o data frame somente com as linhas que contém os status direcionado e em tratamento
    df = df[(df['Tipo de Produto'] != 'PORTAL DE NEGOCIOS') & (df['Subcategoria'] != 'PORTAL DE NEGOCIOS - AUTO') & (df['Subcategoria'] != 'PORTAL DE NEGOCIOS - RE')]#ignora os incidentes classificados como portal de negócios. Esses INs são direcionados manualmente

    
    #Le as tabelas de de_para
    df_desc = pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\DE_PARA_TRIAGEM_DESCRICAO.xlsx', sheet_name='Planilha1')
    df_tipo = pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\DE_PARA_TRIAGEM_TIPO_PRODUTO.xlsx', sheet_name='Planilha1')

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
                if (descricao.find(palavra) != -1):
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
    df.to_excel(triagem_arquivo_final, 'Planilha1',index=False)#gera o arquivo de destino sem a coluna de indice que é gerada automaticamente pelo dataframe do pandas
    lista_grp_destino = df['GRUPO_DESTINO'].values
    total_incidentes = len(lista_grp_destino) #total de incidentes na lista
    indeterminados = np.count_nonzero(lista_grp_destino == "INDETERMINADO")
    redirecionaveis = total_incidentes - indeterminados #indica quantos incidentes devem ser direcionados
    print("Resumo da classificação de incidentes:")
    print(str(total_incidentes) + " incidentes no grupo de triagem")
    print(str(redirecionaveis) + " incidentes que serão direcionados automaticamente")
    print(str(indeterminados) + " incidentes que não puderam ser classificados pelo robô. Eles precisam ser direcionados manualmente.")
    print("========================================Fim da classificação========================================")

def acessa_pesquisa_incidentes():
    driver = instancia_driver()
    driver.get('https://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')#acessa a pagina do SIGS
    login(driver)#executa o login
    painel_esquerda(driver, 'Gerenciamento de Incidentes', 'Pesquisar Incidentes', 3)
    return driver

def pesquisa_incidente(driver, id_incidente, segunda_execucao):
    time.sleep(4)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')#obtem a lista dos iframes
    if (segunda_execucao == False):
        IS.troca_frame(driver, lista_objetos[1])#seleciona o iframe do formulario de pesquisa do incidente
        IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X20'), id_incidente)#preenche o id do incidente
        driver.switch_to.default_content()#retorna ao content default
        IS.clica_xpath(driver, '//button[text()="Pesquisar"]')#clica em pesquisar
    else:#caso seja a segunda execução ou posterior
        IS.retorna_objetos(driver, 'id', 'X20').clear() #apaga o id
        IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X20'), id_incidente)#preenche o id do incidente
        driver.switch_to.default_content()#retorna ao content default
        IS.clica_xpath(driver, '//button[text()="Pesquisar"]')#clica em pesquisar

def triagem_tipifica(driver, id, tipo_produto, descricao_resumida, descricao):
    df_tipo_produto_abend = pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\TIPO_PRODUTO_ABEND.xlsx', sheet_name='Planilha1')
    contem_tp_produto_abend = df_tipo_produto_abend['TIPO_PRODUTO'].str.contains(tipo_produto).any()
    if (contem_tp_produto_abend == True):
        if ('APLICACAO' in descricao_resumida):#se tem 'aplicação' no título do incidente é disfunção
            tipificacao = 'DISFUNÇÃO'
        else:#se tiver o S222 na descrição é disfunção
            if ('S222' in descricao):
                tipificacao = 'DISFUNÇÃO'
            else:
                tipificacao = 'INCIDENTE (ABEND/INTERRUPÇÃO)'    
    else:    
        tipificacao = 'DISFUNÇÃO'
    IS.retorna_objetos(driver, 'id', 'X321').clear()
    time.sleep(1)
    text_area = IS.retorna_objetos(driver, 'id', 'X321')
    time.sleep(1)
    IS.insere_texto(text_area, tipificacao)#seleciona a tipificação
    time.sleep(1)

def trata_excecao_janela_salvar(driver):
    try:
        botao = IS.retorna_objetos(driver,'xpath','//button[text()="Não"]')
        print(botao.get_attribute("id"))

    except:
        #print('botão não encontrado')
        driver.switch_to.default_content()#retorna ao content default

    try:
        IS.clica_xpath_time(driver, '//button[text()="Não"]', 3)
        #print('clicou no botão não')
    except:
        #print('O botão "não" não existe, o robô pode continuar')
        pass

def redireciona_incidente(driver, id_incidente, grupo_destino, expande, tipo_produto, tipificacao, descricao_resumida, descricao):
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    IS.troca_frame(driver, lista_objetos[1])#troca para o frame do formulario do incidente
    time.sleep(4)
    IS.retorna_objetos(driver, 'id', 'X35').clear()
    IS.retorna_objetos(driver, 'id', 'X35').send_keys(grupo_destino)
    if(expande == True):#se a aba de atividades ainda não foi expandida
        IS.clica_xpath(driver, '//span[text()="Atividades"]')#clica para expandir a aba
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X261'), 'CORRIGIR DIRECIONAMENTO')#insere o texto "Corrigir Direcionamento"
    time.sleep(1)
    text_area = IS.retorna_objetos(driver, 'id', 'X272View')
    time.sleep(1)
    IS.insere_texto(text_area, 'Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente...')#insere a justificativa do redirecionamento
    time.sleep(1)
    #se a tipificação inicial está vazia, tipificar agora
    if (type(tipificacao)==np.float64):#float é porque o dado vem como NAN. NAN é um nulo float no pandas
        triagem_tipifica(driver, id_incidente, tipo_produto, descricao_resumida, descricao)
    time.sleep(1)
    driver.switch_to.default_content()#retorna ao content default
    IS.clica_xpath(driver, '//button[text()="Salvar"]')#salva
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'xpath', '//button[text()="OK"]')
    try :
        IS.clica_id_time(driver, 'o',3)
    except:
        pass
    time.sleep(2)
    trata_excecao_janela_salvar(driver)#chama a função que trata a exceção da janela de salvar
    IS.clica_xpath_time(driver, '//button[text()="Cancelar"]', 4)#clica para sair do incidente
    time.sleep(2)
    trata_excecao_janela_salvar(driver)
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    IS.troca_frame(driver, lista_objetos[-2])#troca para o frame do formulario do incidente

def inicia_redirecionamento():
    lista_ids = [] #lista que armazena os ids dos incidentes
    lista_grp_destino = [] #lista que são armazenados os grupos de destino
    lista_tipificacao = [] #lista que receberá as tipificações dos incidentes
    df = pd.read_excel(triagem_arquivo_final, sheet_name = 'Planilha1') 
    lista_ids = df['ID do Incidente'].values 
    lista_grp_destino = df['GRUPO_DESTINO'].values
    lista_tipificacao = df['Brd Tp in'].values
    expande = True #variavel de controle para determinar se a aba "Atividades" no incidente precisa ser expandida ou não 
    segunda_execucao = False #variável de controle para impedir que haja estouro de indice da lista de iframes a partir da segunda execução
    incidentes_redirecionados = 0 #contador de incidentes direcionados
    total_incidentes = len(lista_grp_destino) #total de incidentes na lista
    indeterminados = np.count_nonzero(lista_grp_destino == "INDETERMINADO")
    redirecionaveis = total_incidentes - indeterminados #indica quantos incidentes devem ser direcionados
    if (redirecionaveis > 0):
        print("========================================Início do redirecionamento automático========================================")
        driver = acessa_pesquisa_incidentes()
        for x in range(total_incidentes):
            if(lista_grp_destino[x] != 'INDETERMINADO'):
                pesquisa_incidente(driver, lista_ids[x], segunda_execucao)#pesquisa o incidente
                redireciona_incidente(driver, lista_ids[x], lista_grp_destino[x], expande, df['Tipo de Produto'][x], lista_tipificacao[x], df['Descrição Resumida'][x], df['Descrição'][x] )#chama a função que redireciona o incidente passando o id
                incidentes_redirecionados += 1
                print("Incidentes direcionados: " + str(incidentes_redirecionados) + " de " + str(redirecionaveis))        
                expande = False #depois da primeira execução não é mais necessário clicar na aba para expandí-la
                segunda_execucao = True #depois da primeira execução é necessário jogar para True
        logoff(driver)#faz logoff e encerra o driver
        print("========================================Fim do redirecionamento automático========================================")


def main_triagem():
    main_extrai_incidentes(2)#extrai os incidentes
    triagem_classifica()
    inicia_redirecionamento() #o driver recebe 0 quando não houve redirecionamento
    

def tipifica_incidente(driver, id, expande, tipificacao):#driver, id do incidente, flag que determina se a aba 'atividade' será precisa ser expandida
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    IS.troca_frame(driver, lista_objetos[1])#troca para o frame do formulario do incidente
    time.sleep(4)
    if(expande == True):#se a aba de atividades ainda não foi expandida
        IS.clica_xpath(driver, '//span[text()="Atividades"]')#clica para expandir a aba
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X261'), 'CORRIGIR DIRECIONAMENTO')#insere o texto "Corrigir Direcionamento"
    time.sleep(1)
    text_area = IS.retorna_objetos(driver, 'id', 'X272View')
    time.sleep(1)
    IS.insere_texto(text_area, 'Tipificando incidente... Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...')#insere a justificativa do redirecionamento
    time.sleep(1)
    IS.retorna_objetos(driver, 'id', 'X322').clear()
    time.sleep(1)
    text_area = IS.retorna_objetos(driver, 'id', 'X322')
    time.sleep(1)
    try:
        IS.insere_texto(text_area, tipificacao)#seleciona a tipificação
    except:
        pass
    time.sleep(1)
    driver.switch_to.default_content()#retorna ao content default
    IS.clica_xpath(driver, '//button[text()="Salvar"]')#salva
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'xpath', '//button[text()="OK"]')
    try:
        IS.clica_id_time(driver, 'o',3)
    except:
        pass
    time.sleep(2)
    trata_excecao_janela_salvar(driver)#chama a função que trata a exceção da janela de salvar
    time.sleep(1)
    IS.clica_xpath_time(driver, '//button[text()="Cancelar"]', 2)#clica para sair do incidente
    time.sleep(2)
    trata_excecao_janela_salvar(driver)
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    IS.troca_frame(driver, lista_objetos[-2])#troca para o frame do formulario do incidente

def inicia_tipificacao():
    lista_ids = [] #lista que armazena os ids dos incidentes
    df_incidentes =  pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\Dashboard Incidentes\\base dashboard incidentes.xlsx', sheet_name='Planilha1')
    vazio = [np.nan, np.nan]#variável para facilitar a obtenção dos incidentes que a tipificação esteja vazia
    df_incidentes = df_incidentes[df_incidentes['Brd Tp in'].isin(vazio)]#seleciona somente os incidentes que não estão tipificados
    status = ['DIRECIONADO']#lista de status aceitos no arquivo
    df_incidentes = df_incidentes[df_incidentes['Status'].isin(status)]

    df_tipo_produto_abend = pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\TIPO_PRODUTO_ABEND.xlsx', sheet_name='Planilha1')
    #df_incidentes = df_incidentes[~df_incidentes['Tipo de Produto'].isin(df_tipo_produto_abend['TIPO_PRODUTO'])]

    lista_ids = df_incidentes['ID do Incidente'].values 
    lista_tipo_produto = df_incidentes['Tipo de Produto'].values 
    expande = True #variavel de controle para determinar se a aba "Atividades" no incidente precisa ser expandida ou não 
    segunda_execucao = False #variável de controle para impedir que haja estouro de indice da lista de iframes a partir da segunda execução
    incidentes_tipificados = 0 #contador de incidentes tipificados
    total_incidentes = len(lista_ids) #total de incidentes na lista
    tipificacao = ''
    
    print('Incidentes a serem tipificados: ' + str(total_incidentes))
    if (total_incidentes > 0):
        print("========================================Início da tipificação automática========================================")
        driver = acessa_pesquisa_incidentes()
        for x in range(total_incidentes):
            contem_tp_produto_abend = df_tipo_produto_abend['TIPO_PRODUTO'].str.contains(lista_tipo_produto[x]).any()
            if (contem_tp_produto_abend == True):
                descricao_resumida = df_incidentes['Descrição Resumida'].iloc[x]
                if ('APLICACAO' in descricao_resumida):#se tem 'aplicação' no título do incidente é disfunção
                    tipificacao = 'DISFUNÇÃO'
                else:#se tiver o S222 na descrição é disfunção
                    descricao = df_incidentes['Descrição'].iloc[x]
                    if ('S222' in descricao):
                        tipificacao = 'DISFUNÇÃO'
                    else:
                        tipificacao = 'INCIDENTE (ABEND/INTERRUPÇÃO)'    
            else:    
                tipificacao = 'DISFUNÇÃO'
            pesquisa_incidente(driver, lista_ids[x], segunda_execucao)#pesquisa o incidente
            tipifica_incidente(driver, lista_ids[x], expande, tipificacao)#chama a função que tipifica o incidente passando o id
            incidentes_tipificados += 1
            print("Incidentes tipificados: " + str(incidentes_tipificados) + " de " + str(total_incidentes))        
            expande = False #depois da primeira execução não é mais necessário clicar na aba para expandí-la
            segunda_execucao = True #depois da primeira execução é necessário jogar para True
        logoff(driver)#faz logoff e encerra o driver
        print("========================================Fim da tipificação automática========================================")
        

def main_tipifica_incidentes():
    main_extrai_incidentes(3)
    inicia_tipificacao()
    print("Fim da execução=======================")
    

def pesquisa_horas_trabalhadas(driver, grupo, inicio_atividade, indice_frame):#iframe do formulario de pesquisa
    time.sleep(3)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')#obtem a lista dos iframes
    IS.troca_frame(driver, lista_objetos[indice_frame])#seleciona o iframe do formulario de pesquisa do incidente
    time.sleep(2)
    IS.retorna_objetos(driver, 'id', 'X5').clear() #apaga o grupo
    IS.retorna_objetos(driver, 'id', 'X7').clear() #apaga o inicio da atividade
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X5'), grupo)#preenche o grupo
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X7'), inicio_atividade)#preenche o inicio da atividade
    driver.switch_to.default_content()#retorna ao content default
    time.sleep(2)
    IS.clica_xpath(driver, '//button[text()="Pesquisar"]')#clica em pesquisar

def acessa_artefato_sigs(menu, submenu, timeout):
    driver = instancia_driver()
    driver.get('https://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')#acessa a pagina do SIGS
    login(driver)#executa o login
    painel_esquerda(driver, menu, submenu, timeout)
    return driver

def baixa_horas_trabalhadas_grupo(driver, grupo, inicio_atividade, arquivo, cabecalho):
    time.sleep(1)#remover depois
    if (cabecalho==True):
        pesquisa_horas_trabalhadas(driver, grupo, inicio_atividade, -2)#-2 no índice frame porque é o frame a ser utilizado na hora de pesquisar os incidentes na primeira execução. -3 na segunda execução em diante
    else:
        pesquisa_horas_trabalhadas(driver, grupo, inicio_atividade, -3)#-2 no índice frame porque é o frame a ser utilizado na hora de pesquisar os incidentes na primeira execução. -3 na segunda execução em diante
    pagina_lista_artefatos_pesquisados(driver)
    pagina_exportacao(driver, arquivo, cabecalho, 200)
    return driver

def retorna_pesquisa(driver):
    time.sleep(3)
    driver.switch_to.default_content()#retorna ao content default
    time.sleep(3)
    IS.clica_xpath(driver, '//button[text()="Voltar"]')#clica em voltar

def deleta_arquivos_temporarios(filenames, arquivo_merge):
    #deleta os arquivos temporários
    for arquivo in filenames:#remove os arquivos baixados do sigs
        if os.path.exists(arquivo):
            os.remove(arquivo)
    if os.path.exists(arquivo_merge):#remove o arquivo concatenado
        os.remove(arquivo_merge)

def main_horas_trabalhadas():
    df_grupos =  pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_HORAS_TRABALHADAS.xlsx', sheet_name='Plan1')
    df_grupos = df_grupos[df_grupos['ATIVO?'] == 'S']#remove os grupos inativos

    #montar uma lista com os arquivos que serão criados
    filenames = ['//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export.txt']
    for indice in range(len(df_grupos['GRUPOS']) - 1):#ignora o nome do arquivo do primeiro grupo, pois o primeiro foi inserido manualmente na lista filenames
        filenames.append('//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export ('+  str(indice+1) + ').txt')

    #declaração dos arquivos que serão manipulados
    arquivo_merge = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/arq.txt' #arquivo resultante do merge
    arquivo_final = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/HORAS_TRABALHADAS/horas_trabalhadas.xlsx' #arquivo convertido em xlsx
    
    deleta_arquivos_temporarios(filenames, arquivo_merge)#deleta arquivos temporários caso haja algo no diretório

    cabecalho = True #flag para definir se é a primeira execução. caso tenha mais de um arquivo baixado, é necessário passar a flag para o método de download para que desmarque a opção de baixar o arquivo com cabeçalho das colunas
    


    #calculo do dia do ano passado para reduzir a volumetria das extrações
    hoje = date.today()
    um_ano = timedelta(367)
    dia_ano_passado = str((hoje - um_ano).strftime("%d/%m/%y")) + ' 23:59:59'

    driver = acessa_artefato_sigs('', 'Consulta de Horas Trabalhadas', 4)

    for x in range(len(filenames)):
        baixa_horas_trabalhadas_grupo(driver, df_grupos['GRUPOS'][x] , dia_ano_passado, filenames[x], cabecalho)
        cabecalho = False
        retorna_pesquisa(driver)
    
    logoff(driver)

    #faz o merge dos N arquivos baixados
    with open(arquivo_merge,  'w', encoding='utf-8') as outfile:#cria o arquivo de saída
        for fname in filenames:
            with open(fname, encoding='utf-8') as infile:
                for line in infile:
                    outfile.write(line)#escreve cada linha no arquivo de destino
    df_horas = pd.read_csv(arquivo_merge, encoding='utf-8', sep='\t') #o arquivo tem o encoding ansi, então é necessário marcar isso 
                                                            #juntamente com o delimitador sep='\t' que significa por tab

    #converte os campos de data para o excel
    df_horas["Data Inicio Servico"] = pd.to_datetime(df_horas["Data Inicio Servico"], format='%d/%m/%Y %H:%M:%S')#converte os data types para datetime
    df_horas["Data Fim Servico"] = pd.to_datetime(df_horas["Data Fim Servico"], format='%d/%m/%Y %H:%M:%S')#converte os data types para datetime

    #cria e preenche os campos novos
    df_horas['tempo_segundos'] = '' #cria a coluna de segundos
    df_horas['RS'] = '' #cria a coluna que receberá o nome da RS
    for indice in range (len(df_horas['Tempo Atividade'])):#preenche a coluna de segundos
        tamanho = len(df_horas['Tempo Atividade'][indice])
        tamanho_substr = tamanho - 8
        tempo = df_horas['Tempo Atividade'][indice]
        horas_segundos = int(tempo[tamanho_substr:tamanho_substr+2])*3600
        minutos_segundos = int(tempo[tamanho_substr+3:tamanho_substr+5])*60
        segundos = int(tempo[-2:])
        total_segundos = horas_segundos + minutos_segundos + segundos
    
        if (tamanho_substr > 0):
                dias_segundos = int(tempo[0:tamanho_substr])*24*3600
                total_segundos+=dias_segundos
        df_horas['tempo_segundos'][indice] = total_segundos
    for indice in range(len(df_horas['Número'])):#preenche a coluna com os nomes das RSs
        artefato = df_horas['Número'][indice]
        if ('-' in artefato):
            artefato = artefato[:artefato.find('-')]
            df_horas['RS'][indice] = artefato
    
    df_horas.to_excel(arquivo_final, 'Planilha1',index=False)#cria o arquivo xlsx
    deleta_arquivos_temporarios(filenames, arquivo_merge)#deleta os arquivos temporários
    print('Arquivo de horas trabalhadas gerado')

def gera_historico_incidentes(tipo_historico):
    if tipo_historico==1: main_extrai_incidentes(tipo_historico)#código 1 significa que é necessário baixar os incidentes novamente antes de criar o histórico
    
    now = datetime.now()

    hora = now.strftime("%Y-%m-%d %H:%M:%S")
    ano_mes_agora = str(now.strftime("%Y%m"))
    print("Hora de agora: ", hora)
    
    #declara os arquivos
    arq_historico = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/teste_historico_ins.xlsx'
    arq_extracao = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx'

    #cria os dataframes
    db_connection_str = 'mysql+pymysql://root:admin@localhost/teste'
    db_connection = create_engine(db_connection_str, encoding='utf-8')
    df_historico = pd.read_sql('SELECT ID_INCIDENTE, STATUS, GRUPO_SIGS, DATA_ABERTURA, DATA_RESOLUCAO, DATA_EXTRACAO FROM historico_incidente', con=db_connection)#lê a tabela de histórico no banco de dados
    df_extracao = pd.read_excel(arq_extracao)#dataframe do arquivo de extração do SIGS
    df_insert = pd.DataFrame(columns=['ID_INCIDENTE', 'STATUS', 'GRUPO_SIGS','DATA_ABERTURA','DATA_RESOLUCAO', 'DATA_EXTRACAO'])#dataframe utilizado no insert da tabela de histórico

    #lista status
    lista_status = ['DIRECIONADO', 'EM TRATAMENTO']

    print('Tamanho histórico: {}'.format(len(df_historico)))

    tamanho_df_extracao = len(df_extracao['ID do Incidente'])

    print('Tamanho df extração: {}'.format(tamanho_df_extracao))
    
    #contadores
    appends = 0
    ins_novos = 0
    atualizacoes = 0
    transbordo_prox_mes = 0

    #for indice in range(len(df_extracao['ID do Incidente'])):
    for indice in range(tamanho_df_extracao):#Busca o incidente na lista histórica
        lista_historico = df_historico[df_historico['ID_INCIDENTE']==df_extracao['ID do Incidente'][indice]]
        if (len(lista_historico)>0):#se achou o IN no histórico então carrega
            designacao_principal = lista_historico['GRUPO_SIGS'].tail(1).values[0]
            status = lista_historico['STATUS'].tail(1).values[0]
            if ((df_extracao['Designação principal'][indice]!=designacao_principal) | (df_extracao['Status'][indice]!=status)):#carrega se o grupo ou o status mudarem
                linha = list([df_extracao['ID do Incidente'][indice], df_extracao['Status'][indice], df_extracao['Designação principal'][indice], df_extracao['Hora de Abertura'][indice] ,df_extracao['Hora de Resolução'][indice],hora])
                df_linha = pd.DataFrame([linha], columns=['ID_INCIDENTE', 'STATUS', 'GRUPO_SIGS','DATA_ABERTURA','DATA_RESOLUCAO', 'DATA_EXTRACAO'])
                df_insert = pd.concat([df_insert, df_linha], ignore_index=True)
                appends +=1
                atualizacoes +=1
                #print(df_historico.tail(1))
            elif ((status in (lista_status))):
                data_extracao = str(lista_historico['DATA_EXTRACAO'].tail(1).values[0])
                ano_mes_extracao = str(data_extracao[0:4]) + str(data_extracao[5:7])
                if(ano_mes_agora!=ano_mes_extracao):#se ano mês de hoje for diferente do ano mes do IN no histórico, carrega
                    print('Ano mês agora: {}'.format(ano_mes_agora))
                    print('Ano mês extração: {}'.format(ano_mes_extracao))
                    linha = list([df_extracao['ID do Incidente'][indice], df_extracao['Status'][indice], df_extracao['Designação principal'][indice], df_extracao['Hora de Abertura'][indice] ,df_extracao['Hora de Resolução'][indice],hora])
                    df_linha = pd.DataFrame([linha], columns=['ID_INCIDENTE', 'STATUS', 'GRUPO_SIGS','DATA_ABERTURA','DATA_RESOLUCAO', 'DATA_EXTRACAO'])
                    df_insert = pd.concat([df_insert, df_linha], ignore_index=True)
                    appends +=1
                    transbordo_prox_mes +=1
        else:#não achou o IN no histórico, então carrega
            linha = list([df_extracao['ID do Incidente'][indice], df_extracao['Status'][indice], df_extracao['Designação principal'][indice], df_extracao['Hora de Abertura'][indice] ,df_extracao['Hora de Resolução'][indice],hora])
            df_linha = pd.DataFrame([linha], columns=['ID_INCIDENTE', 'STATUS', 'GRUPO_SIGS','DATA_ABERTURA','DATA_RESOLUCAO', 'DATA_EXTRACAO'])
            df_insert = pd.concat([df_insert, df_linha], ignore_index=True)
            appends +=1
            ins_novos +=1
            #df_historico.tail(1)
    
    df_historico = pd.concat([df_historico, df_insert], ignore_index=True)
    df_insert.to_sql('historico_incidente', con=db_connection, if_exists='append', index=False)


    print('Appends: {}'.format(appends))
    print('INs novos: {}'.format(ins_novos))
    print('Atualizações: {}'.format(atualizacoes))
    print('Transbordo para o outro mês: {}'.format(transbordo_prox_mes))

    print('Tamanho histórico: {}'.format(len(df_historico)))
    print('Tamanho df extração: {}'.format(tamanho_df_extracao))

    #converte para data
    df_historico["DATA_ABERTURA"] = pd.to_datetime(df_historico["DATA_ABERTURA"], format='%Y-%m-%d %H:%M:%S')#converte os data types para datetime
    df_historico["DATA_RESOLUCAO"] = pd.to_datetime(df_historico["DATA_RESOLUCAO"], format='%Y-%m-%d %H:%M:%S')#converte os data types para datetime
    df_historico["DATA_EXTRACAO"] = pd.to_datetime(df_historico["DATA_EXTRACAO"], format='%Y-%m-%d %H:%M:%S')#converte os data types para datetime
 
    #carrega o arquivo de histórico
    df_historico.to_excel(arq_historico, 'Planilha1',index=False)#gera o arquivo de destino sem a coluna de indice que é gerada automaticamente pelo dataframe do pandas

def busca_incidente_historico(incidente):
    return incidente
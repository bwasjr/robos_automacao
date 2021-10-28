import numpy as np
import time
import os
import glob
import pandas as pd
import csv
from xlsxwriter.workbook import Workbook
from statistics import mode
import interacoes_selenium_alfa as IS  

#declaração dos arquivos que serão manipulados
filenames = ["//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/export.txt"] #arquivos baixados pelo robo
arquivo_merge = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/arq.txt' #arquivo resultante do merge
arquivo_final = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/extracao_robo_sust_bare.xlsx' #arquivo convertido em xlsx

def login(driver):
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'LoginUsername'), 'g571602')
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'LoginPassword'), 'was7822')
    IS.clica_id(driver, 'loginBtn')

def painel_esquerda(driver):#selecoes do painel da equerda
    try:
        IS.clica_id(driver, 'o') #clica na janela automática
    except:
        pass
    IS.clica_classe(driver, 'x-panel-header', 4)#expande o gerenciar incidentes
    IS.clica_id(driver, 'ROOT/Gerenciamento de Incidentes/Pesquisar Incidentes') #clica em pesquisar incidentes

def pesquisa_incidentes(driver, grupo, aberto_apos):#iframe do formulario de pesquisa
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')#obtem a lista dos iframes
    IS.troca_frame(driver, lista_objetos[-2])#seleciona o iframe do formulario de pesquisa do incidente
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X114'), grupo)#preenche o grupo 'DS - BS - SUSTENTACAO-BARE*'
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X31'), aberto_apos)#preenche o aberto após '31/12/2018 23:59:59'
    driver.switch_to.default_content()#retorna ao content default
    IS.clica_xpath(driver, '//button[text()="Pesquisar"]')#clica em pesquisar

def pagina_lista_incidentes(driver):#pagina que contém a lista dos incidentes que serao exportados
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'button')
    IS.clica_por_texto(lista_objetos,'Mais') #clica no botao 'mais'
    lista_objetos.clear()
    lista_objetos = IS.retorna_objetos(driver, 'class', 'x-menu-item-text') #retorna os botoes apos clicar em mais
    IS.clica_objeto_lista(lista_objetos,2)# clica em exportar para arquivo de texto

def acessa_pesquisa_incidentes():
    driver = IS.instancia_driver()
    driver.get('https://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')#acessa a pagina do SIGS
    login(driver)#executa o login
    painel_esquerda(driver)
    return driver

def pagina_exportacao(driver, arquivo):#frame da pagina de exportacao
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')#obtem a lista de iframes
    IS.troca_frame(driver, lista_objetos[1])#clica no frame do radio buttons
    lista_objetos.clear()
    IS.clica_id(driver, 'X10Label')#clica no tabulação
    IS.clica_id(driver, 'X21')#clica no botao verde de ok

def main_run(grupo, aberto_apos, arquivo):
    driver = acessa_pesquisa_incidentes()
    pesquisa_incidentes(driver, grupo, aberto_apos)
    pagina_lista_incidentes(driver)
    pagina_exportacao(driver, arquivo)
    return driver

def logoff(driver):
    driver.switch_to.default_content()
    lista_objetos = IS.retorna_objetos(driver, 'class','logout')
    lista_objetos[0].click()
    alerta = driver.switch_to.alert
    time.sleep(1)
    alerta.accept()
    time.sleep(1)
    driver.quit()

def extrai_incidentes():
    print("========================================\nInício da extração\n========================================")
    driver = main_run('DS - BS - SUSTENTACAO-BARE', '31/12/2018 23:59:59', filenames[0])
    logoff(driver)
    driver.quit()

def classifica():
    print("========================================\nInício da classificação\n========================================")
    with open(arquivo_merge, 'w') as outfile:#cria o arquivo de saída
        for fname in filenames:
            with open(fname) as infile:
                for line in infile:
                    outfile.write(line)#escreve cada linha no arquivo de destino

    df = pd.read_csv(arquivo_merge, encoding='ansi', sep='\t') #o arquivo tem o encoding ansi, então é necessário marcar isso juntamente com o delimitador sep='\t' que significa por tab
    df = df[df['Designação principal'] == 'DS - BS - SUSTENTACAO-BARE'] #remove as linhas que não tenham o grupo DS - BS - SUSTENTACAO-BARE
    status = ['DIRECIONADO']#lista de status aceitos no arquivo
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
    df.to_excel(arquivo_final, 'Planilha1',index=False)#gera o arquivo de destino sem a coluna de indice que é gerada automaticamente pelo dataframe do pandas

    for arquivo in filenames:#remove os baixados do sigs
        if os.path.exists(arquivo):
            os.remove(arquivo)

    if os.path.exists(arquivo_merge):#remove o arquivo concatenado
        os.remove(arquivo_merge)

def inicia_redirecionamento():
    print("========================================\nInício do redirecionamento\n========================================")
    driver = acessa_pesquisa_incidentes()
    lista_ids = [] #lista que armazena os ids dos incidentes
    lista_grp_destino = [] #lista que são armazenados os grupos de destino
    df = pd.read_excel(arquivo_final, shee_tname = 'Planilha1') 
    lista_ids = df['ID do Incidente'].values 
    lista_grp_destino = df['GRUPO_DESTINO'].values

    expande = True #variavel de controle para determinar se a aba "Atividades" no incidente precisa ser expandida ou não 
    segunda_execucao = False #variável de controle para impedir que haja estouro de indice da lista de iframes a partir da segunda execução
    incidentes_redirecionados = 0 #contador de incidentes direcionados

    for x in range(len(lista_grp_destino)):
        if(lista_grp_destino[x] != 'INDETERMINADO'):
            pesquisa_incidente(driver, lista_ids[x], segunda_execucao)#pesquisa o incidente
            redireciona_incidente(driver, lista_ids[x], lista_grp_destino[x], expande)#chama a função que redireciona o incidente passando o id
            incidentes_redirecionados += 1
            print("Incidentes direcionados: " + str(incidentes_redirecionados))        
            expande = False #depois da primeira execução não é mais necessário clicar na aba para expandí-la
            segunda_execucao = True #depois da primeira execução é necessário jogar para True
    logoff(driver)#faz logoff e encerra o driver    

def pesquisa_incidente(driver, id_incidente, segunda_execucao):
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

def redireciona_incidente(driver, id_incidente, grupo_destino, expande):
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    IS.troca_frame(driver, lista_objetos[1])#troca para o frame do formulario do incidente
    IS.retorna_objetos(driver, 'id', 'X36').clear()
    IS.retorna_objetos(driver, 'id', 'X36').send_keys(grupo_destino)
    if(expande == True):#se a aba de atividades ainda não foi expandida
        IS.clica_xpath(driver, '//span[text()="Atividades"]')#clica para expandir a aba
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X243'), 'CORRIGIR DIRECIONAMENTO')#insere o texto "Corrigir Direcionamento"
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X254'), 'TESTE ROBÔ SIGS')#insere a justificativa do redirecionamento
    driver.switch_to.default_content()#retorna ao content default
    IS.clica_xpath(driver, '//button[text()="Salvar"]')#salva
    lista_objetos = IS.retorna_objetos(driver, 'xpath', '//button[text()="OK"]')
    try :
        IS.clica_id(driver, 'o')
    except:
        print('janela de ok não encontrada')
    
    IS.clica_xpath(driver, '//button[text()="Cancelar"]')#clica para sair do incidente
    
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    IS.troca_frame(driver, lista_objetos[-2])#troca para o frame do formulario do incidente
    try:
        IS.clica_id(driver, 'n') #clica no botão 
    except:
        print('Botão "Não" não encontrado')

#main run
extrai_incidentes()
classifica()
inicia_redirecionamento()
print("=======================\nFim da execução\n=======================")

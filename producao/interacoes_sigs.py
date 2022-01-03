from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
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
# import pymysql
import datetime
import sys
from dateutil.relativedelta import *


# arquivos globais
# arquivo convertido em xlsx
triagem_arquivo_entrada = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/extracao_robo.xlsx'
# arquivo com os incidentes classificados
triagem_arquivo_final = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/classificao_triagem.xlsx'

inicio_execucao = time.perf_counter()


def gera_df_grupos(tipo_execucao: int) -> pd.DataFrame:
    if tipo_execucao == 1:  # faz extração de uma lista completa: grupos da sustentação + projeto + SAP
        df_grupos = pd.read_excel(
            '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_EXTRACAO.xlsx', sheet_name='Plan1')
    elif tipo_execucao == 2:  # extração para triagem
        df_grupos = pd.read_excel(
            '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_EXTRACAO_TRIAGEM.xlsx', sheet_name='Plan1')
    elif tipo_execucao == 3:  # extração para tipificação dos incidentes
        df_grupos = pd.read_excel(
            '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_EXTRACAO_TIPIFICACAO.xlsx', sheet_name='Plan1')
    df_grupos = df_grupos[df_grupos['ATIVO?']
                          == 'S']  # remove os grupos inativos
    return df_grupos


def gera_lista_extracao(df_grupos: pd.DataFrame) -> list:
    # montar uma lista com os arquivos que serão criados
    filenames = [
        '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export.txt']
    # ignora o nome do arquivo do primeiro grupo, pois o primeiro foi inserido manualmente na lista filenames
    for indice in range(len(df_grupos['GRUPOS']) - 1):
        filenames.append(
            '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export (' + str(indice+1) + ').txt')
    return filenames


def get_dia_ano_passado() -> str:
    hoje = date.today()
    um_ano = timedelta(367)
    dia_ano_passado = str((hoje - um_ano).strftime("%d/%m/%y")) + ' 23:59:59'
    return dia_ano_passado


def merge_arquivos_baixados(arquivo_merge, filenames):
    with open(arquivo_merge,  'w', encoding='utf-8') as outfile:  # cria o arquivo de saída
        for fname in filenames:
            with open(fname, encoding='utf-8') as infile:
                for line in infile:
                    # escreve cada linha no arquivo de destino
                    outfile.write(line)
    return arquivo_merge


def converte_df_datatypes(df: pd.DataFrame) -> pd.DataFrame:
    df["Hora de Abertura"] = pd.to_datetime(
        df["Hora de Abertura"], format='%d/%m/%Y %H:%M:%S')
    df["Hora de Atualização"] = pd.to_datetime(
        df["Hora de Atualização"], format='%d/%m/%Y %H:%M:%S')  # converte os data types para datetime
    df["Hora de Resolução"] = pd.to_datetime(
        df["Hora de Resolução"], format='%d/%m/%Y %H:%M:%S')  # converte os data types para datetime
    df["Hora de Fechamento"] = pd.to_datetime(
        df["Hora de Fechamento"], format='%d/%m/%Y %H:%M:%S')  # converte os data types para datetime
    df["Hora de Reabertura"] = pd.to_datetime(
        df["Hora de Reabertura"], format='%d/%m/%Y %H:%M:%S')  # converte os data types para datetime
    # converte os data types para datetime
    df["Hora do Alerta"] = pd.to_datetime(
        df["Hora do Alerta"], format='%d/%m/%Y %H:%M:%S')
    return df


def gera_arquivos_secundarios():
    gera_historico_incidentes()
    gera_historico_estoque()
    gera_incidentes_removidos()
    gera_indice_resolucao(0)  # gera o arquivo geral
    # gera o indice de resolução dos incidentes do portal
    gera_indice_resolucao(1)


def instancia_driver():  # INICIALIZACAO do chromedriver
    options = webdriver.ChromeOptions()
    # options.headless = True
    # options.add_argument('--headless')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument("--test-type")
    options.add_argument("--start-maximized")
    prefs = {'download.prompt_for_download': False, 'download.default_directory':
             '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\downloads', 'download.directory_upgrade': True}
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(
        'C:\\Users\\g571602\\Documents\\Python\\robo_bare\\chromedriver.exe', options=options)
    return driver


def login(driver):
    # coloque a matrícula do usuário
    IS.insere_texto(IS.retorna_objetos(
        driver, 'id', 'LoginUsername'), 'g571602')
    IS.insere_texto(IS.retorna_objetos(
        driver, 'id', 'LoginPassword'), 'was78122')  # coloque a senha
    IS.clica_id(driver, 'loginBtn')


def logoff(driver):
    time.sleep(1)
    driver.switch_to.default_content()
    IS.clica_id(driver, 'toolbarUserInfoButtonId')
    lista_objetos = IS.retorna_objetos(driver, 'class', 'icon-user-logout')
    lista_objetos[0].click()
    alerta = driver.switch_to.alert
    alerta.accept()
    driver.quit()  # encerra o driver


# selecoes do painel da equerda e interage com o artefato desejado do SIGS
def painel_esquerda(driver, menu, submenu, timeout):
    time.sleep(2)
    try:
        IS.clica_id_time(driver, 'o', timeout)  # clica na janela automática
    except:
        pass
    time.sleep(3)
    # obtem a lista de abas para clicar
    lista_objetos = IS.retorna_objetos(driver, 'class', 'x-panel-header')
    if (menu != ''):
        # clica no item de menu desejado de acordo com o texto
        IS.clica_por_texto_time(lista_objetos, menu, timeout)
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'class', 'x-tree-node-el')
    # clica no submenu desejado
    IS.clica_por_texto_time(lista_objetos, submenu, timeout)


# iframe do formulario de pesquisa
def pesquisa_incidentes(driver, grupo, aberto_apos, indice_frame, tipo_execucao):
    time.sleep(3)
    lista_objetos = IS.retorna_objetos(
        driver, 'tag', 'iframe')  # obtem a lista dos iframes
    # seleciona o iframe do formulario de pesquisa do incidente
    IS.troca_frame(driver, lista_objetos[indice_frame])
    time.sleep(2)
    if (tipo_execucao != 1):
        # Só baixa os incidentes em aberto
        IS.clica_id_time(driver, 'X13Label', 10)
    else:
        # caso contrário a opção "Ambos" deve ser clicada e incidentes abertos e encerrados serão baixados
        IS.clica_id_time(driver, 'X17Label', 10)
    IS.retorna_objetos(driver, 'id', 'X114').clear()  # apaga o grupo
    IS.retorna_objetos(driver, 'id', 'X31').clear()  # apaga o aberto apos
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X114'),
                    grupo)  # preenche o grupo
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X31'),
                    aberto_apos)  # preenche o aberto após
    driver.switch_to.default_content()  # retorna ao content default
    time.sleep(2)
    # clica em pesquisar
    IS.clica_xpath(driver, '//button[text()="Pesquisar"]')


# pagina que contém a lista dos incidentes que serao exportados
def pagina_lista_artefatos_pesquisados(driver):
    time.sleep(8)
    lista_objetos = IS.retorna_objetos(driver, 'xpath', '//button')
    time.sleep(1)
    IS.clica_por_texto(lista_objetos, 'Mais')  # clica no botao 'mais'
    lista_objetos.clear()
    # retorna os botoes apos clicar em mais
    lista_objetos = IS.retorna_objetos(driver, 'class', 'x-menu-item-text')
    # clica em exportar para arquivo de texto
    IS.clica_objeto_lista(lista_objetos, -2)


def aguarda_download(timeout):
    print('aguardando o download')
    count = 0  # variavel de controle do timeout do download do arquivo
    arquivo = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export.txt'
    while (not os.path.exists(arquivo) and count < timeout):
        time.sleep(1)
        count += 1
        if (count == timeout):
            print('O download do arquivo excedeu o timeout de ' +
                  str(timeout) + ' segundos')
            return 'timeout'
    if (count < timeout):
        print('O arquivo foi baixado em ' + str(count) + ' segundos')
        return 'ok'


def pagina_exportacao(driver, cabecalho, timeout):  # frame da pagina de exportacao
    time.sleep(3)
    lista_objetos = IS.retorna_objetos(
        driver, 'tag', 'iframe')  # obtem a lista de iframes
    # clica no frame do radio buttons
    IS.troca_frame(driver, lista_objetos[-1])
    lista_objetos.clear()
    lista_objetos = IS.retorna_objetos(
        driver, 'tag', 'label')  # obtem a lista de iframes
    if cabecalho == False:
        # clica para remover o cabeçalho das colunas
        IS.clica_id(driver, 'X2Edit')
    IS.clica_id(driver, 'X10Label')  # clica no tabulação
    IS.clica_id(driver, 'X21')  # clica no botao verde de ok
    is_timeout = aguarda_download(timeout)
    if is_timeout == 'timeout':
        return 'timeout'
    else:
        return 'ok'


def second_run_incidentes(driver, grupo, aberto_apos, cabecalho, tipo_execucao, timeout_arquivo, numero_arquivo):
    time.sleep(8)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'button')
    time.sleep(1)
    driver.switch_to.default_content()  # retorna ao content default
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'xpath', '//button')
    time.sleep(1)
    IS.clica_por_texto(lista_objetos, 'Voltar')  # clica em voltar
    # -3 é o íncide do iframe. Deve ser -3 na segunda execução em diante
    pesquisa_incidentes(driver, grupo, aberto_apos, -3, tipo_execucao)
    pagina_lista_artefatos_pesquisados(driver)
    is_timeout = pagina_exportacao(driver, cabecalho, timeout_arquivo)
    if is_timeout == 'ok':
        nome_arquivo = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export (' + str(
            numero_arquivo) + ').txt'
        os.rename('//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export.txt', nome_arquivo)
        # renomeia o arquivo baixado para evitar a sobrescreção na execução headless


def baixa_incidentes_grupo(grupo, aberto_apos, arquivo, cabecalho, tipo_execucao):
    driver = instancia_driver()
    # acessa a pagina do SIGS
    driver.get('https://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')
    login(driver)  # executa o login
    painel_esquerda(driver, 'Gerenciamento de Incidentes',
                    'Pesquisar Incidentes', 4)
    time.sleep(1)  # remover depois
    # -2 no índice frame porque é o frame a ser utilizado na hora de pesquisar os incidentes na primeira execução. -3 na segunda execução em diante
    pesquisa_incidentes(driver, grupo, aberto_apos, -2, tipo_execucao)
    pagina_lista_artefatos_pesquisados(driver)
    pagina_exportacao(driver, cabecalho, 530)
    return driver


# principal execução responsável por extrair os incidentes
def main_extrai_incidentes(tipo_execucao: int):
    df_grupos = gera_df_grupos(tipo_execucao)
    filenames = gera_lista_extracao(df_grupos)

    # declaração dos arquivos que serão manipulados
    arquivo_merge = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/arq.txt'
    # arquivo convertido em xlsx
    arquivo_final = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx'
    arquivo_extracao_robo = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/extracao_robo.xlsx'

    deleta_arquivos_temporarios(filenames, arquivo_merge)
    cabecalho = True  # flag para definir se o cabeçalho do arquivo será baixado
    # dia do ano passado para reduzir a volumetria das extrações
    dia_ano_passado = get_dia_ano_passado()

    # extrai o primeiro arquivo
    driver = baixa_incidentes_grupo(
        df_grupos['GRUPOS'][0], dia_ano_passado, filenames[0], cabecalho, tipo_execucao)
    # renomeia o primeiro arquivo para evitar que a execução headless sobrescreva este arquivo
    os.rename('//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export.txt',
              '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export_com_cabecalho.txt')
    cabecalho = False  # depois da primeira execução o cabeçalho não é mais necessário
    # inicia a execução do segundo arquivo em diante
    for index in range(len(filenames)-1):
        second_run_incidentes(
            driver, df_grupos['GRUPOS'][index+1], dia_ano_passado, cabecalho, tipo_execucao, 600, index+1)
    time.sleep(2)
    # renomeia o primeiro arquivo para o nome original a fim de facilitar o merge e utilizá-lo primeiro para aproveitar seu cabeçalho
    os.rename('//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export_com_cabecalho.txt',
              '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export.txt')
    logoff(driver)
    arquivo_merge = merge_arquivos_baixados(arquivo_merge, filenames)
    # o arquivo tem o encoding ansi, então é necessário converter
    df = pd.read_csv(arquivo_merge, encoding='utf-8', sep='\t')
    df = converte_df_datatypes(df)  # converte os data types para datetime

    # remoção de aspas no id do incidente
    df['ID do Incidente'].replace('"', '', inplace=True)
    df['ID do Incidente'].replace('\n', '', inplace=True)
    df['ID do Incidente'].replace('\t', '', inplace=True)

    # cria a coluna de portal de negócios
    df['is_portal'] = ''
    binario = (df['Tipo de Produto'] == 'PORTAL DE NEGOCIOS') | (df['Subcategoria'] == 'PORTAL DE NEGOCIOS - AUTO') | (df['Subcategoria']
                                                                                                                       == 'PORTAL DE NEGOCIOS - RE')  # ignora os incidentes classificados como portal de negócios. Esses INs são direcionados manualmente
    df['is_portal'] = binario
    if tipo_execucao in [1, 3]:
        # gera o arquivo de destino sem a coluna de indice que é gerada automaticamente pelo dataframe do pandas
        df.to_excel(arquivo_final, 'Planilha1', index=False)
        # gera o arquivo de entrada pra triagem do robô
        df.to_excel(arquivo_extracao_robo, 'Planilha1', index=False)
    elif tipo_execucao == 2:
        # gera o arquivo de entrada pra triagem do robô
        df.to_excel(arquivo_extracao_robo, 'Planilha1', index=False)
    deleta_arquivos_temporarios(filenames, arquivo_merge)

    print('Fim da extração de incidentes')
    calcula_tempo_execucao()
    driver.quit()


def triagem_classifica():
    print("========================================Início da classificação========================================")
    # o arquivo tem o encoding ansi, então é necessário marcar isso juntamente com o delimitador sep='\t' que significa por tab
    df = pd.read_excel(triagem_arquivo_entrada)
    # remove as linhas que não tenham o grupo DS - BS - SUSTENTACAO-BARE
    df = df[df['Designação principal'] == 'DS - BS - SUSTENTACAO-BARE']
    # faz triagem somente dos incidentes de produção. Os de outros ambientes não são tratados pela sustentação
    df = df[df['Brd Tipo Ambiente'] == 'PRODUCAO']
    status = ['DIRECIONADO']  # lista de status aceitos no arquivo
    # sustitui o data frame somente com as linhas que contém os status direcionado e em tratamento
    df = df[df['Status'].isin(status)]
    df = df[(df['Tipo de Produto'] != 'PORTAL DE NEGOCIOS') & (df['Subcategoria'] != 'PORTAL DE NEGOCIOS - AUTO') & (df['Subcategoria']
                                                                                                                     != 'PORTAL DE NEGOCIOS - RE')]  # ignora os incidentes classificados como portal de negócios. Esses INs são direcionados manualmente

    # Le as tabelas de de_para
    df_desc = pd.read_excel(
        '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\DE_PARA_TRIAGEM_DESCRICAO.xlsx', sheet_name='Planilha1')
    df_tipo = pd.read_excel(
        '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\DE_PARA_TRIAGEM_TIPO_PRODUTO.xlsx', sheet_name='Planilha1')

    lista_grupos = []
    count = -1  # contador para acompanhar o indice do dataframe

    for tipo in df['Tipo de Produto']:  # pega o tipo de produto da lista de incidentes
        count += 1
        # se o tipo de produto da lista de incidentes existe na tabela de de_para
        if (tipo in df_tipo['TIPO_PRODUTO'].values):
            # recupera o indice do tipo de produto na depara
            tipo_depara_index = df_tipo[df_tipo['TIPO_PRODUTO']
                                        == tipo].index.values
            # recupera o nome do grupo na depara
            grupo = df_tipo['GRUPO'][tipo_depara_index]
            # adiciona o nome do grupo na lista de grupos de destino
            lista_grupos.append(grupo.values[0])
        else:  # tenta classificar com a descrição
            # busca o a descricao do incidente. foi passada a linha com o count e a coluna com o 8
            descricao = df.iloc[count, 8]
            lista_matches = []  # lista para armazenar todos os codigos de grupos que ocorreram matches entre palavras e descricoes
            for palavra in df_desc['PALAVRAS']:
                # pega o indice da palavra no arquivo de DE_PARA
                palavra_index = df_desc[df_desc['PALAVRAS']
                                        == palavra].index.values
                if (descricao.find(palavra) != -1):
                    # pega o código do grupo de destino correspondente à palavra
                    cd_grupo = int(df_desc['CODIGO_GRUPO'][palavra_index])
                    # adiciona codigo do grupo na lista de matches
                    lista_matches.append(cd_grupo)
            # elege o grupo mais votado
            primeiro = max(set(lista_matches),
                           key=lista_matches.count, default=-1)
            # pega o indice da palavra no arquivo de DE_PARA
            grupo_index = df_desc[df_desc['CODIGO_GRUPO']
                                  == primeiro].index.values
            if (len(grupo_index) > 0):  # se houve matches
                # acessa o nome do grupo pela primeira ocorrencia de indice na lista grupo_index
                grupo = df_desc['GRUPO_DESTINO'][grupo_index[0]]
                lista_grupos.append(grupo)
            else:  # se não foram encontradas correspondências de palavras
                lista_grupos.append('INDETERMINADO')
    # cria a coluna dos grupos de destino no dataframe
    df['GRUPO_DESTINO'] = lista_grupos
    # gera o arquivo de destino sem a coluna de indice que é gerada automaticamente pelo dataframe do pandas
    df.to_excel(triagem_arquivo_final, 'Planilha1', index=False)
    lista_grp_destino = df['GRUPO_DESTINO'].values
    total_incidentes = len(lista_grp_destino)  # total de incidentes na lista
    indeterminados = np.count_nonzero(lista_grp_destino == "INDETERMINADO")
    # indica quantos incidentes devem ser direcionados
    redirecionaveis = total_incidentes - indeterminados
    print("Resumo da classificação de incidentes:")
    print(str(total_incidentes) + " incidentes no grupo de triagem")
    print(str(redirecionaveis) +
          " incidentes que serão direcionados automaticamente")
    print(str(indeterminados) +
          " incidentes que não puderam ser classificados pelo robô. Eles precisam ser direcionados manualmente.")
    print("========================================Fim da classificação========================================")


def acessa_pesquisa_incidentes():
    driver = instancia_driver()
    # acessa a pagina do SIGS
    driver.get('https://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')
    login(driver)  # executa o login
    painel_esquerda(driver, 'Gerenciamento de Incidentes',
                    'Pesquisar Incidentes', 3)
    return driver


def pesquisa_incidente(driver, id_incidente, segunda_execucao):
    time.sleep(4)
    lista_objetos = IS.retorna_objetos(
        driver, 'tag', 'iframe')  # obtem a lista dos iframes
    if (segunda_execucao == False):
        # seleciona o iframe do formulario de pesquisa do incidente
        IS.troca_frame(driver, lista_objetos[1])
        IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X20'),
                        id_incidente)  # preenche o id do incidente
        driver.switch_to.default_content()  # retorna ao content default
        # clica em pesquisar
        IS.clica_xpath(driver, '//button[text()="Pesquisar"]')
    else:  # caso seja a segunda execução ou posterior
        IS.retorna_objetos(driver, 'id', 'X20').clear()  # apaga o id
        IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X20'),
                        id_incidente)  # preenche o id do incidente
        driver.switch_to.default_content()  # retorna ao content default
        # clica em pesquisar
        IS.clica_xpath(driver, '//button[text()="Pesquisar"]')


def triagem_tipifica(driver, id, tipo_produto, descricao_resumida, descricao):
    df_tipo_produto_abend = pd.read_excel(
        '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\TIPO_PRODUTO_ABEND.xlsx', sheet_name='Planilha1')
    contem_tp_produto_abend = df_tipo_produto_abend['TIPO_PRODUTO'].str.contains(
        tipo_produto).any()
    if (contem_tp_produto_abend == True):
        if ('APLICACAO' in descricao_resumida):  # se tem 'aplicação' no título do incidente é disfunção
            tipificacao = 'DISFUNÇÃO'
        else:  # se tiver o S222 na descrição é disfunção
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
    IS.insere_texto(text_area, tipificacao)  # seleciona a tipificação
    time.sleep(1)


def trata_excecao_janela_salvar(driver):
    try:
        botao = IS.retorna_objetos(driver, 'xpath', '//button[text()="Não"]')
        print(botao.get_attribute("id"))
    except:
        # print('botão não encontrado')
        driver.switch_to.default_content()  # retorna ao content default
    try:
        # clica no botão para fechar a janela de salvar
        IS.clica_xpath_time(driver, '//button[text()="Não"]', 3)
    except TimeoutException:
        print('A janela de prompt para salvar não surgiu, o robô pode continuar')
        pass


def redireciona_incidente(driver, id_incidente, grupo_destino, expande, tipo_produto, tipificacao, descricao_resumida, descricao):
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    # troca para o frame do formulario do incidente
    IS.troca_frame(driver, lista_objetos[1])
    time.sleep(4)
    IS.retorna_objetos(driver, 'id', 'X36').clear()
    IS.retorna_objetos(driver, 'id', 'X36').send_keys(grupo_destino)
    if(expande == True):  # se a aba de atividades ainda não foi expandida
        # clica para expandir a aba
        IS.clica_xpath(driver, '//span[text()="Atividades"]')
    # insere o texto "Corrigir Direcionamento"
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X262'),
                    'CORRIGIR DIRECIONAMENTO')
    time.sleep(1)
    text_area = IS.retorna_objetos(driver, 'id', 'X273View')
    time.sleep(1)
    IS.insere_texto(text_area, 'Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente... Redirecionando incidente...')  # insere a justificativa do redirecionamento
    time.sleep(1)
    # se a tipificação inicial está vazia, tipificar agora
    # float é porque o dado vem como NAN. NAN é um nulo float no pandas
    if (type(tipificacao) == np.float64):
        triagem_tipifica(driver, id_incidente, tipo_produto,
                         descricao_resumida, descricao)
    time.sleep(1)
    driver.switch_to.default_content()  # retorna ao content default
    IS.clica_xpath(driver, '//button[text()="Salvar"]')  # salva
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(
        driver, 'xpath', '//button[text()="OK"]')
    try:
        IS.clica_id_time(driver, 'o', 3)
    except:
        pass
    time.sleep(2)
    # chama a função que trata a exceção da janela de salvar
    trata_excecao_janela_salvar(driver)
    time.sleep(1)
    # clica para sair do incidente
    IS.clica_xpath_time(driver, '//button[text()="Cancelar"]', 2)
    time.sleep(2)
    trata_excecao_janela_salvar(driver)
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    # troca para o frame do formulario do incidente
    IS.troca_frame(driver, lista_objetos[-2])


def inicia_redirecionamento():
    lista_ids = []  # lista que armazena os ids dos incidentes
    lista_grp_destino = []  # lista que são armazenados os grupos de destino
    lista_tipificacao = []  # lista que receberá as tipificações dos incidentes
    df = pd.read_excel(triagem_arquivo_final, sheet_name='Planilha1')
    lista_ids = df['ID do Incidente'].values
    lista_grp_destino = df['GRUPO_DESTINO'].values
    lista_tipificacao = df['Brd Tp in'].values
    expande = True  # variavel de controle para determinar se a aba "Atividades" no incidente precisa ser expandida ou não
    # variável de controle para impedir que haja estouro de indice da lista de iframes a partir da segunda execução
    segunda_execucao = False
    incidentes_redirecionados = 0  # contador de incidentes direcionados
    total_incidentes = len(lista_grp_destino)  # total de incidentes na lista
    indeterminados = np.count_nonzero(lista_grp_destino == "INDETERMINADO")
    # indica quantos incidentes devem ser direcionados
    redirecionaveis = total_incidentes - indeterminados
    if (redirecionaveis > 0):
        print("========================================Início do redirecionamento automático========================================")
        driver = acessa_pesquisa_incidentes()
        for x in range(total_incidentes):
            if(lista_grp_destino[x] != 'INDETERMINADO'):
                # pesquisa o incidente
                pesquisa_incidente(driver, lista_ids[x], segunda_execucao)
                redireciona_incidente(driver, lista_ids[x], lista_grp_destino[x], expande, df['Tipo de Produto'][x], lista_tipificacao[x],
                                      df['Descrição Resumida'][x], df['Descrição'][x])  # chama a função que redireciona o incidente passando o id
                incidentes_redirecionados += 1
                print("Incidentes direcionados: " +
                      str(incidentes_redirecionados) + " de " + str(redirecionaveis))
                # depois da primeira execução não é mais necessário clicar na aba para expandí-la
                expande = False
                segunda_execucao = True  # depois da primeira execução é necessário jogar para True
        logoff(driver)  # faz logoff e encerra o driver
        print("========================================Fim do redirecionamento automático========================================")


def main_triagem():
    main_extrai_incidentes(2)  # extrai os incidentes
    triagem_classifica()
    inicia_redirecionamento()  # o driver recebe 0 quando não houve redirecionamento


# driver, id do incidente, flag que determina se a aba 'atividade' será precisa ser expandida
def tipifica_incidente(driver, id, expande, tipificacao):
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    # troca para o frame do formulario do incidente
    IS.troca_frame(driver, lista_objetos[1])
    time.sleep(4)
    if(expande == True):  # se a aba de atividades ainda não foi expandida
        # clica para expandir a aba
        IS.clica_xpath(driver, '//span[text()="Atividades"]')
    # insere o texto "Corrigir Direcionamento"
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X262'),
                    'CORRIGIR DIRECIONAMENTO')
    time.sleep(1)
    text_area = IS.retorna_objetos(driver, 'id', 'X273View')
    time.sleep(1)
    IS.insere_texto(text_area, 'Tipificando incidente... Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...  Tipificando incidente...')  # insere a justificativa do redirecionamento
    time.sleep(1)
    IS.retorna_objetos(driver, 'id', 'X323').clear()
    time.sleep(1)
    text_area = IS.retorna_objetos(driver, 'id', 'X323')
    time.sleep(1)
    try:
        IS.insere_texto(text_area, tipificacao)  # seleciona a tipificação
    except:
        pass
    time.sleep(1)
    driver.switch_to.default_content()  # retorna ao content default
    IS.clica_xpath(driver, '//button[text()="Salvar"]')  # salva
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(
        driver, 'xpath', '//button[text()="OK"]')
    try:
        IS.clica_id_time(driver, 'o', 3)
    except:
        pass
    time.sleep(2)
    # chama a função que trata a exceção da janela de salvar
    trata_excecao_janela_salvar(driver)
    time.sleep(1)
    # clica para sair do incidente
    IS.clica_xpath_time(driver, '//button[text()="Cancelar"]', 2)
    time.sleep(2)
    trata_excecao_janela_salvar(driver)
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    # troca para o frame do formulario do incidente
    IS.troca_frame(driver, lista_objetos[-2])


def inicia_tipificacao():
    lista_ids = []  # lista que armazena os ids dos incidentes
    df_incidentes = pd.read_excel(
        '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\Dashboard Incidentes\\base dashboard incidentes.xlsx', sheet_name='Planilha1')
    # variável para facilitar a obtenção dos incidentes que a tipificação esteja vazia
    vazio = [np.nan, np.nan]
    # seleciona somente os incidentes que não estão tipificados
    df_incidentes = df_incidentes[df_incidentes['Brd Tp in'].isin(vazio)]
    status = ['DIRECIONADO']  # lista de status aceitos no arquivo
    df_incidentes = df_incidentes[df_incidentes['Status'].isin(status)]

    df_tipo_produto_abend = pd.read_excel(
        '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\TIPO_PRODUTO_ABEND.xlsx', sheet_name='Planilha1')
    # df_incidentes = df_incidentes[~df_incidentes['Tipo de Produto'].isin(df_tipo_produto_abend['TIPO_PRODUTO'])]

    lista_ids = df_incidentes['ID do Incidente'].values
    lista_tipo_produto = df_incidentes['Tipo de Produto'].values
    expande = True  # variavel de controle para determinar se a aba "Atividades" no incidente precisa ser expandida ou não
    # variável de controle para impedir que haja estouro de indice da lista de iframes a partir da segunda execução
    segunda_execucao = False
    incidentes_tipificados = 0  # contador de incidentes tipificados
    total_incidentes = len(lista_ids)  # total de incidentes na lista
    tipificacao = ''

    print('Incidentes a serem tipificados: ' + str(total_incidentes))
    if (total_incidentes > 0):
        print("========================================Início da tipificação automática========================================")
        driver = acessa_pesquisa_incidentes()
        for x in range(total_incidentes):
            contem_tp_produto_abend = df_tipo_produto_abend['TIPO_PRODUTO'].str.contains(
                lista_tipo_produto[x]).any()
            if (contem_tp_produto_abend == True):
                descricao_resumida = df_incidentes['Descrição Resumida'].iloc[x]
                # se tem 'aplicação' no título do incidente é disfunção
                if ('APLICACAO' in descricao_resumida):
                    tipificacao = 'DISFUNÇÃO'
                else:  # se tiver o S222 na descrição é disfunção
                    descricao = df_incidentes['Descrição'].iloc[x]
                    if ('S222' in descricao):
                        tipificacao = 'DISFUNÇÃO'
                    else:
                        tipificacao = 'INCIDENTE (ABEND/INTERRUPÇÃO)'
            else:
                tipificacao = 'DISFUNÇÃO'
            # pesquisa o incidente
            pesquisa_incidente(driver, lista_ids[x], segunda_execucao)
            # chama a função que tipifica o incidente passando o id
            tipifica_incidente(driver, lista_ids[x], expande, tipificacao)
            incidentes_tipificados += 1
            print("Incidentes tipificados: " +
                  str(incidentes_tipificados) + " de " + str(total_incidentes))
            # depois da primeira execução não é mais necessário clicar na aba para expandí-la
            expande = False
            segunda_execucao = True  # depois da primeira execução é necessário jogar para True
        logoff(driver)  # faz logoff e encerra o driver


def main_tipifica_incidentes():
    main_extrai_incidentes(3)
    inicia_tipificacao()
    print("Fim da tipificação de incidentes=======================")


# iframe do formulario de pesquisa
def pesquisa_horas_trabalhadas(driver, grupo, inicio_atividade, indice_frame):
    time.sleep(3)
    lista_objetos = IS.retorna_objetos(
        driver, 'tag', 'iframe')  # obtem a lista dos iframes
    # seleciona o iframe do formulario de pesquisa do incidente
    IS.troca_frame(driver, lista_objetos[indice_frame])
    time.sleep(2)
    IS.retorna_objetos(driver, 'id', 'X5').clear()  # apaga o grupo
    # apaga o inicio da atividade
    IS.retorna_objetos(driver, 'id', 'X7').clear()
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X5'),
                    grupo)  # preenche o grupo
    IS.insere_texto(IS.retorna_objetos(driver, 'id', 'X7'),
                    inicio_atividade)  # preenche o inicio da atividade
    driver.switch_to.default_content()  # retorna ao content default
    time.sleep(2)
    # clica em pesquisar
    IS.clica_xpath(driver, '//button[text()="Pesquisar"]')


def acessa_artefato_sigs(menu, submenu, timeout):
    driver = instancia_driver()
    # acessa a pagina do SIGS
    driver.get('https://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')
    login(driver)  # executa o login
    painel_esquerda(driver, menu, submenu, timeout)
    return driver


def baixa_horas_trabalhadas_grupo(driver, grupo, inicio_atividade, arquivo, cabecalho):
    time.sleep(1)  # remover depois
    if (cabecalho == True):
        # -2 no índice frame porque é o frame a ser utilizado na hora de pesquisar os incidentes na primeira execução. -3 na segunda execução em diante
        pesquisa_horas_trabalhadas(driver, grupo, inicio_atividade, -2)
    else:
        # -2 no índice frame porque é o frame a ser utilizado na hora de pesquisar os incidentes na primeira execução. -3 na segunda execução em diante
        pesquisa_horas_trabalhadas(driver, grupo, inicio_atividade, -3)
    pagina_lista_artefatos_pesquisados(driver)
    pagina_exportacao(driver, cabecalho, 200)
    return driver


def retorna_pesquisa(driver):
    time.sleep(3)
    driver.switch_to.default_content()  # retorna ao content default
    time.sleep(3)
    IS.clica_xpath(driver, '//button[text()="Voltar"]')  # clica em voltar


def deleta_arquivos_temporarios(filenames, arquivo_merge):
    # deleta os arquivos temporários
    primeiro_arquivo_baixado = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export_com_cabecalho.txt'

    for arquivo in filenames:  # remove os arquivos baixados do sigs
        if os.path.exists(arquivo):
            os.remove(arquivo)
    if os.path.exists(arquivo_merge):  # remove o arquivo concatenado
        os.remove(arquivo_merge)
    if os.path.exists(primeiro_arquivo_baixado):  # remove o arquivo concatenado
        os.remove(primeiro_arquivo_baixado)


def main_horas_trabalhadas():
    df_grupos = pd.read_excel(
        '\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\LISTA_GRUPOS_HORAS_TRABALHADAS.xlsx', sheet_name='Plan1')
    df_grupos = df_grupos[df_grupos['ATIVO?']
                          == 'S']  # remove os grupos inativos

    # montar uma lista com os arquivos que serão criados
    filenames = [
        '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export.txt']
    # ignora o nome do arquivo do primeiro grupo, pois o primeiro foi inserido manualmente na lista filenames
    for indice in range(len(df_grupos['GRUPOS']) - 1):
        filenames.append(
            '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/downloads/export (' + str(indice+1) + ').txt')

    # declaração dos arquivos que serão manipulados
    # arquivo resultante do merge
    arquivo_merge = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/robo_sigs/arq.txt'
    # arquivo convertido em xlsx
    arquivo_final = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/HORAS_TRABALHADAS/horas_trabalhadas.xlsx'

    # deleta arquivos temporários caso haja algo no diretório
    deleta_arquivos_temporarios(filenames, arquivo_merge)

    cabecalho = True  # flag para definir se é a primeira execução. caso tenha mais de um arquivo baixado, é necessário passar a flag para o método de download para que desmarque a opção de baixar o arquivo com cabeçalho das colunas

    # calculo do dia do ano passado para reduzir a volumetria das extrações
    hoje = date.today()
    um_ano = timedelta(367)
    dia_ano_passado = str((hoje - um_ano).strftime("%d/%m/%y")) + ' 23:59:59'

    driver = acessa_artefato_sigs('', 'Consulta de Horas Trabalhadas', 4)

    for x in range(len(filenames)):
        baixa_horas_trabalhadas_grupo(
            driver, df_grupos['GRUPOS'][x], dia_ano_passado, filenames[x], cabecalho)
        cabecalho = False
        retorna_pesquisa(driver)

    logoff(driver)

    # faz o merge dos N arquivos baixados
    with open(arquivo_merge,  'w', encoding='utf-8') as outfile:  # cria o arquivo de saída
        for fname in filenames:
            with open(fname, encoding='utf-8') as infile:
                for line in infile:
                    # escreve cada linha no arquivo de destino
                    outfile.write(line)
    # o arquivo tem o encoding ansi, então é necessário marcar isso
    df_horas = pd.read_csv(arquivo_merge, encoding='utf-8', sep='\t')
    # juntamente com o delimitador sep='\t' que significa por tab

    # converte os campos de data para o excel
    df_horas["Data Inicio Servico"] = pd.to_datetime(
        df_horas["Data Inicio Servico"], format='%d/%m/%Y %H:%M:%S')  # converte os data types para datetime
    df_horas["Data Fim Servico"] = pd.to_datetime(
        df_horas["Data Fim Servico"], format='%d/%m/%Y %H:%M:%S')  # converte os data types para datetime

    # cria e preenche os campos novos
    df_horas['tempo_segundos'] = ''  # cria a coluna de segundos
    df_horas['RS'] = ''  # cria a coluna que receberá o nome da RS
    # preenche a coluna de segundos
    for indice in range(len(df_horas['Tempo Atividade'])):
        tamanho = len(df_horas['Tempo Atividade'][indice])
        tamanho_substr = tamanho - 8
        tempo = df_horas['Tempo Atividade'][indice]
        horas_segundos = int(tempo[tamanho_substr:tamanho_substr+2])*3600
        minutos_segundos = int(tempo[tamanho_substr+3:tamanho_substr+5])*60
        segundos = int(tempo[-2:])
        total_segundos = horas_segundos + minutos_segundos + segundos

        if (tamanho_substr > 0):
            dias_segundos = int(tempo[0:tamanho_substr])*24*3600
            total_segundos += dias_segundos
        df_horas['tempo_segundos'][indice] = total_segundos
    # preenche a coluna com os nomes das RSs
    for indice in range(len(df_horas['Número'])):
        artefato = df_horas['Número'][indice]
        if ('-' in artefato):
            artefato = artefato[:artefato.find('-')]
            df_horas['RS'][indice] = artefato

    df_horas.to_excel(arquivo_final, 'Planilha1',
                      index=False)  # cria o arquivo xlsx
    # deleta os arquivos temporários
    deleta_arquivos_temporarios(filenames, arquivo_merge)
    print('Arquivo de horas trabalhadas gerado')


def gera_historico_incidentes():
    now = datetime.datetime.now()
    hora = now.strftime("%Y-%m-%d %H:%M:%S")

    # declara os arquivos
    arq_extracao = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx'
    arq_historico_incidentes = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/historico_incidentes.xlsx'

    # cria os dataframes dos arquivos
    # dataframe do arquivo de extração do SIGS
    df_extracao = pd.read_excel(arq_extracao)
    # dataframe com o histórico dos estoques
    df_historico_incidentes = pd.read_excel(arq_historico_incidentes)

    # cria o dataframe de INs novos
    df_ins_novos = df_extracao[df_extracao['ID do Incidente'].isin(
        df_historico_incidentes['ID_INCIDENTE']) == False]
    # seleciona somente as colunas necessárias para o histórico de incidentes
    df_ins_novos = df_ins_novos[['ID do Incidente', 'Status',
                                 'Designação principal', 'Hora de Abertura', 'Hora de Resolução']]
    df_ins_novos = df_ins_novos.rename(columns={"ID do Incidente": "ID_INCIDENTE", "Status": "STATUS",
                                       "Designação principal": "GRUPO_SIGS", "Hora de Abertura": "DATA_ABERTURA", "Hora de Resolução": "DATA_RESOLUCAO"})
    df_ins_novos['DATA_EXTRACAO'] = hora  # cria a coluna de data de extração

    # cria o dataframe de INs antigos
    df_ins_antigos = df_extracao[df_extracao['ID do Incidente'].isin(
        df_historico_incidentes['ID_INCIDENTE']) == True]
    # seleciona somente as colunas necessárias para o histórico de incidentes
    df_ins_antigos = df_ins_antigos[['ID do Incidente', 'Status',
                                     'Designação principal', 'Hora de Abertura', 'Hora de Resolução']]
    df_ins_antigos = df_ins_antigos.rename(columns={"ID do Incidente": "ID_INCIDENTE", "Status": "STATUS",
                                           "Designação principal": "GRUPO_SIGS", "Hora de Abertura": "DATA_ABERTURA", "Hora de Resolução": "DATA_RESOLUCAO"})

    df_insere = pd.DataFrame(columns=['ID_INCIDENTE', 'STATUS', 'GRUPO_SIGS', 'DATA_ABERTURA',
                             'DATA_RESOLUCAO', 'DATA_EXTRACAO'])  # transforma a linha em um dataframe

    # prepara o dataframe de inserção dos Ins antigos que mudaram de grupo ou status
    for indice in df_ins_antigos.index:
        incidente = df_ins_antigos['ID_INCIDENTE'][indice]
        df_incidente = df_historico_incidentes[df_historico_incidentes['ID_INCIDENTE'].isin([
                                                                                            incidente])]
        if ((df_ins_antigos['STATUS'][indice] != df_incidente['STATUS'].iloc[-1]) | (df_ins_antigos['GRUPO_SIGS'][indice] != df_incidente['GRUPO_SIGS'].iloc[-1])):
            linha = list([df_ins_antigos['ID_INCIDENTE'][indice], df_ins_antigos['STATUS'][indice], df_ins_antigos['GRUPO_SIGS'][indice],
                         df_ins_antigos['DATA_ABERTURA'][indice], df_ins_antigos['DATA_RESOLUCAO'][indice], hora])  # gera a linha do registro
            df_linha = pd.DataFrame([linha], columns=['ID_INCIDENTE', 'STATUS', 'GRUPO_SIGS',
                                    'DATA_ABERTURA', 'DATA_RESOLUCAO', 'DATA_EXTRACAO'])  # transforma a linha em um dataframe
            # insere no dataframe o registro
            df_insere = pd.concat([df_insere, df_linha], ignore_index=True)

    # concatena os ins novos no df_insere
    df_insere = pd.concat([df_insere, df_ins_novos], ignore_index=True)

    # concatena o df insere com o df historico incidentes
    df_historico_incidentes = pd.concat(
        [df_historico_incidentes, df_insere], ignore_index=True)

    # converte a coluna para datetime
    df_historico_incidentes["DATA_EXTRACAO"] = pd.to_datetime(
        df_historico_incidentes["DATA_EXTRACAO"], format='%Y-%m-%d %H:%M:%S')  # converte os data types para datetime

    # gera o novo arquivo do histórico de incidentes
    df_historico_incidentes.to_excel(
        arq_historico_incidentes, 'Planilha1', index=False)

    print('Arquivo de histórico de incidentes gerado')


def gera_historico_estoque():
    now = datetime.datetime.now()

    hora = now.strftime("%Y-%m-%d %H:%M:%S")
    ano_mes_agora = str(now.strftime("%Y%m"))

    # declara os arquivos
    arq_extracao = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx'
    arq_estoque = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/estoque_incidentes.xlsx'
    arq_historico_estoque = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/historico_estoque_incidentes.xlsx'

    # cria os dataframes
    # dataframe do arquivo de extração do SIGS
    df_extracao = pd.read_excel(arq_extracao)
    # dataframe utilizado no insert da tabela de estoque de incidentes
    df_estoque = pd.DataFrame(columns=['ID_INCIDENTE', 'ANO_MES_PENDENTE'])
    # dataframe com o histórico dos estoques
    df_historico_estoque = pd.read_excel(arq_historico_estoque)

    # adiciona colunas no dataframe de extracao
    df_extracao['ANO_MES_ABERTURA'] = ''
    df_extracao['ANO_MES_RESOLUCAO'] = ''
    df_extracao['DATA_EXTRACAO'] = hora
    df_extracao['ANO_MES_EXTACAO'] = ano_mes_agora

    # populando as colunas com o ano_mes
    for indice in range(len(df_extracao)):
        df_extracao['ANO_MES_ABERTURA'][indice] = datetime.datetime.strftime(
            df_extracao['Hora de Abertura'][indice], "%Y%m")
        try:  # devido a um bug com esta linha de código tive que colocar este try, se der erro permanece nulo
            df_extracao['ANO_MES_RESOLUCAO'][indice] = datetime.strftime(
                df_extracao['Hora de Resolução'][indice], "%Y%m")
        except:
            pass
        # carregando o dataframe de estoque
        if ((df_extracao['ANO_MES_RESOLUCAO'][indice] > df_extracao['ANO_MES_ABERTURA'][indice]) | ((df_extracao['ANO_MES_RESOLUCAO'][indice] == '') & (df_extracao['ANO_MES_EXTACAO'][indice] > df_extracao['ANO_MES_ABERTURA'][indice]))):
            df_estoque = gera_estoque(df_extracao['ID do Incidente'][indice], df_extracao['Hora de Abertura']
                                      [indice], df_extracao['Hora de Resolução'][indice], df_extracao['DATA_EXTRACAO'][indice], df_estoque)

    # gera o delta entre o historico de estoque e o estoque atual
    df_historico_estoque['INCIDENTE_ANO_MES_ESTOQUE'] = df_historico_estoque['ID_INCIDENTE'] + \
        df_historico_estoque['ANO_MES_PENDENTE'].map(str)
    df_estoque['INCIDENTE_ANO_MES_ESTOQUE'] = df_estoque['ID_INCIDENTE'] + \
        df_estoque['ANO_MES_PENDENTE'].map(str)

    df_estoque = df_estoque[df_estoque['INCIDENTE_ANO_MES_ESTOQUE'].isin(
        df_historico_estoque['INCIDENTE_ANO_MES_ESTOQUE']) == False]  # gera o delta entre extração e histórico
    # incluí os registros novos no histórico
    df_historico_estoque = pd.concat(
        [df_historico_estoque, df_estoque], ignore_index=True)

    # dropa as colunas desnecessárias
    df_estoque = df_estoque[['ID_INCIDENTE', 'ANO_MES_PENDENTE']]
    df_historico_estoque = df_historico_estoque[[
        'ID_INCIDENTE', 'ANO_MES_PENDENTE']]

    # insere o df_estoque na tabela stage do mysql
    df_estoque.to_excel(arq_estoque, 'Planilha1', index=False)
    df_historico_estoque.to_excel(
        arq_historico_estoque, 'Planilha1', index=False)

    print('fim da geração do histórico de estoque')


def gera_estoque(ID_INCIDENTE, DATA_ABERTURA, DATA_RESOLUCAO, DATA_EXTRACAO, df_estoque):
    DATA_ABERTURA = str(DATA_ABERTURA)
    DATA_RESOLUCAO = str(DATA_RESOLUCAO)
    DATA_EXTRACAO = str(DATA_EXTRACAO)
    ano_inicial = int(DATA_ABERTURA[0:4])
    mes_inicial = int(DATA_ABERTURA[5:7])
    try:
        ano_final = int(DATA_RESOLUCAO[0:4])
        mes_final = int(DATA_RESOLUCAO[5:7])
        qt_meses_estoque = (ano_final - ano_inicial) * \
            12 + (mes_final - mes_inicial)
    except:
        ano_final = int(DATA_EXTRACAO[0:4])
        mes_final = int(DATA_EXTRACAO[5:7])
        qt_meses_estoque = (ano_final - ano_inicial) * \
            12 + (mes_final - mes_inicial)

    # criar uma linha no dataframe para cada mês de estoque
    for x in range(qt_meses_estoque):
        data_futura = datetime.datetime(
            year=ano_inicial, month=mes_inicial, day=1)
        data_futura = data_futura + relativedelta(months=+x)
        # gera o ano_mes que o incidente ficou no estoque
        ano_mes_futuro = datetime.datetime.strftime(data_futura, "%Y%m")
        # gera a linha do registro
        linha = list([ID_INCIDENTE, ano_mes_futuro])
        # transforma a linha em um dataframe
        df_linha = pd.DataFrame(
            [linha], columns=['ID_INCIDENTE', 'ANO_MES_PENDENTE'])
        # insere no dataframe o registro
        df_estoque = pd.concat([df_estoque, df_linha], ignore_index=True)
    return df_estoque


def gera_incidentes_removidos():
    now = datetime.datetime.now()
    hora = now.strftime("%Y-%m-%d %H:%M:%S")

    # declara os arquivos
    arq_extracao = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx'
    arq_historico_incidentes = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/historico_incidentes.xlsx'
    arq_incidentes_removidos = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/incidentes_removidos.xlsx'

    # cria os dataframes dos arquivos
    # dataframe do arquivo de extração do SIGS
    df_extracao = pd.read_excel(arq_extracao)
    # dataframe com o histórico dos estoques
    df_historico_incidentes = pd.read_excel(arq_historico_incidentes)
    df_incidentes_removidos = pd.read_excel(arq_incidentes_removidos)

    # aplica filtros
    df_aux_removidos = df_historico_incidentes[df_historico_incidentes['ID_INCIDENTE'].isin(
        df_extracao['ID do Incidente']) == False]

    # retorna o id dos registros agrupados por id_incidente cuja data_extração é a maior
    idx = df_aux_removidos.groupby(['ID_INCIDENTE'])['DATA_EXTRACAO'].transform(
        max) == df_aux_removidos['DATA_EXTRACAO']
    df_aux_removidos = df_aux_removidos[idx]
    df_aux_removidos = df_aux_removidos[df_aux_removidos['STATUS'].isin(
        ['ENCERRADO']) == False]

    # remove as colunas desnecessárias
    df_aux_removidos = df_aux_removidos.drop(['DATA_RESOLUCAO'], axis=1)

    df_aux_removidos = df_aux_removidos.rename(
        columns={"STATUS": "ULTIMO_STATUS_BARE", "GRUPO_SIGS": "ULTIMO_GRUPO_BARE", "DATA_EXTRACAO": "DATA_ULTIMA_EXTRACAO"})  # renomeia as colunas
    df_aux_removidos['DATA_VERIFICACAO_SAIDA'] = hora  # cria coluna
    df_aux_removidos['DATA_VERIFICACAO_SAIDA'] = pd.to_datetime(
        df_aux_removidos['DATA_VERIFICACAO_SAIDA'], format='%Y-%m-%d %H:%M:%S')  # converte os data types para datetime

    lista_indices = []
    df_aux_removidos.reset_index(drop=True, inplace=True)

    for indice in df_aux_removidos.index:
        incidente = df_aux_removidos['ID_INCIDENTE'][indice]
        ultima_extracao = df_aux_removidos['DATA_ULTIMA_EXTRACAO'][indice]
        # if ((incidente in df_incidentes_removidos['ID_INCIDENTE'].values) & (ultima_extracao in df_incidentes_removidos['DATA_ULTIMA_EXTRACAO'].values)==False | (incidente in df_incidentes_removidos['ID_INCIDENTE'].values)==False):
        if (((incidente in df_incidentes_removidos['ID_INCIDENTE'].values) & (ultima_extracao in df_incidentes_removidos['DATA_ULTIMA_EXTRACAO'].values) == False) | ((incidente in df_incidentes_removidos['ID_INCIDENTE'].values) == False)):
            lista_indices.append(indice)

    df_aux_removidos = df_aux_removidos[df_aux_removidos.index.isin(
        lista_indices)]

    print('len aux removidos final ' + str(len(df_aux_removidos)))

    print('len df removidos ' + str(len(df_incidentes_removidos)))

    df_incidentes_removidos = pd.concat(
        [df_incidentes_removidos, df_aux_removidos], ignore_index=True)

    print('len df removidos após concat ' + str(len(df_incidentes_removidos)))

    df_incidentes_removidos.to_excel(
        arq_incidentes_removidos, 'Planilha1', index=False)

    print('Arquivo de incidentes removidos da BARE gerado')


def gera_indice_resolucao(portal):
    # declara os arquivos
    arq_extracao = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/base dashboard incidentes.xlsx'
    if (portal == 1):  # se o parâmetro for portal ==1, então gera o arquivo do portal
        arq_indice_resolucao = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/indice_resolucao_portal.xlsx'
    else:
        arq_indice_resolucao = '//srv-arquivos07/dirgerti/SEDT_AUTORE/TN_AUTORE/SUSTENTAÇÃO/Incidentes/Dashboard Incidentes/indice_resolucao.xlsx'

    # cria os dataframes dos arquivos
    # dataframe do arquivo de extração do SIGS
    df_criados = pd.read_excel(arq_extracao)
    # dataframe do arquivo de extração do SIGS
    df_resolvidos = pd.read_excel(arq_extracao)

    if (portal == 1):  # considera somente os incidentes do portal
        # no excel aparece como VERDADEIRO, mas é porque a variável é boleana e o excel é em portugues
        df_criados = df_criados[df_criados['is_portal'] == True]
        df_resolvidos = df_resolvidos[df_resolvidos['is_portal'] == True]

    # df_resolvidos
    df_resolvidos = df_resolvidos[df_resolvidos['Hora de Resolução'].notna()]
    df_resolvidos['ANO_MES_RESOLUCAO'] = df_resolvidos['Hora de Resolução'].astype(
        str).str[:4]
    df_resolvidos['ANO_MES_RESOLUCAO'] = (df_resolvidos['ANO_MES_RESOLUCAO'].astype(
        int)*100) + df_resolvidos['Hora de Resolução'].astype(str).str[5:7].astype(int)
    df_resolvidos_por_grupo = pd.DataFrame({'count_resolvidos': df_resolvidos.groupby(
        ['ANO_MES_RESOLUCAO', 'Designação principal']).size()}).reset_index()

    # df_criados
    df_criados['ANO_MES_ABERTURA'] = df_criados['Hora de Abertura'].astype(
        str).str[:4]
    df_criados['ANO_MES_ABERTURA'] = (df_criados['ANO_MES_ABERTURA'].astype(
        int)*100) + df_criados['Hora de Abertura'].astype(str).str[5:7].astype(int)
    df_criados = pd.DataFrame({'count': df_criados.groupby(
        ['ANO_MES_ABERTURA', 'Designação principal']).size()}).reset_index()

    # populando o df com os valores
    df_resolvidos_por_grupo['count_criados'] = ''
    lista_contagens = []
    count_for = 0
    for indice in df_resolvidos_por_grupo.index:
        count_for += 1
        ano_mes = df_resolvidos_por_grupo['ANO_MES_RESOLUCAO'].loc[indice]
        grupo = df_resolvidos_por_grupo['Designação principal'].loc[indice]
        aux = df_criados.loc[(df_criados['ANO_MES_ABERTURA'] == ano_mes) & (
            df_criados['Designação principal'] == grupo)]
        if (len(aux) == 0):
            lista_contagens.append(0)
        else:
            count = aux['count'].astype(int)
            for numero in count:
                valor = numero
                lista_contagens.append(valor)

    df_resolvidos_por_grupo['count_criados'] = lista_contagens
    df_resolvidos_por_grupo['indice_resolucao'] = df_resolvidos_por_grupo['count_resolvidos'] / \
        df_resolvidos_por_grupo['count_criados']
    df_resolvidos_por_grupo.to_excel(
        arq_indice_resolucao, index=False)  # gera o arquivo


def calcula_tempo_execucao():
    fim_execucao = time.perf_counter()
    print('O processo total executou em {} segundos'.format(
        round(fim_execucao - inicio_execucao, 0)))

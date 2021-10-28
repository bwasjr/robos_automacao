import csv
import glob
import os
import time
from statistics import mode
import numpy as np
import pandas as pd
from xlsxwriter.workbook import Workbook
import interacoes_selenium as IS
import robo_sigs_extrai_incidentes_v19 as EX
import robo_triagem_v16 as TRI

def inicia_tipificacao():
    lista_ids = [] #lista que armazena os ids dos incidentes
    df_incidentes =  pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\Dashboard Incidentes\\base dashboard incidentes.xlsx', shee_tname='Planilha1')
    vazio = [np.nan, np.nan]#variável para facilitar a obtenção dos incidentes que a tipificação esteja vazia
    df_incidentes = df_incidentes[df_incidentes['Brd Tp in'].isin(vazio)]#seleciona somente os incidentes que não estão tipificados
    status = ['DIRECIONADO']#lista de status aceitos no arquivo
    df_incidentes = df_incidentes[df_incidentes['Status'].isin(status)]

    df_tipo_produto_abend = pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\TIPO_PRODUTO_ABEND.xlsx', shee_tname='Planilha1')
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
        driver = TRI.acessa_pesquisa_incidentes()
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
            TRI.pesquisa_incidente(driver, lista_ids[x], segunda_execucao)#pesquisa o incidente
            tipifica_incidente(driver, lista_ids[x], expande, tipificacao)#chama a função que tipifica o incidente passando o id
            incidentes_tipificados += 1
            print("Incidentes tipificados: " + str(incidentes_tipificados) + " de " + str(total_incidentes))        
            expande = False #depois da primeira execução não é mais necessário clicar na aba para expandí-la
            segunda_execucao = True #depois da primeira execução é necessário jogar para True
        EX.logoff(driver)#faz logoff e encerra o driver
        print("========================================Fim da tipificação automática========================================")
        return driver #terminou a tipificação, então retorna o driver
    return 0 #não ocorreram tipificações , então retorna 0    

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
    IS.retorna_objetos(driver, 'id', 'X321').clear()
    time.sleep(1)
    text_area = IS.retorna_objetos(driver, 'id', 'X321')
    time.sleep(1)
    IS.insere_texto(text_area, tipificacao)#seleciona a tipificação
    time.sleep(1)
    driver.switch_to.default_content()#retorna ao content default
    IS.clica_xpath(driver, '//button[text()="Salvar"]')#salva
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'xpath', '//button[text()="OK"]')
    try :
        IS.clica_time(driver, 'o',3)
    except:
        #print('janela de ok não encontrada')
        pass
    time.sleep(2)
    lista_objetos.clear()
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    try :
        driver.switch_to.frame(lista_objetos[2])
        time.sleep(2)
    except:
        #print('iframe não encontrado')
        driver.switch_to.default_content()#retorna ao content default
    try:
        IS.clica_time(driver, 'n',3)
    except:
        #print('botão de "Deseja salvar?" não encontrado')
        driver.switch_to.default_content()#retorna ao content default
    time.sleep(2)
    IS.clica_xpath(driver, '//button[text()="Cancelar"]')#clica para sair do incidente
    time.sleep(2)
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    IS.troca_frame(driver, lista_objetos[-2])#troca para o frame do formulario do incidente
    try:
        IS.clica_time(driver, 'n',3) #clica no botão 
    except:
        #print('Botão "Não" não encontrado')
        pass

def triagem_tipifica(driver, id, tipo_produto, descricao_resumida, descricao):
    df_tipo_produto_abend = pd.read_excel('\\\\srv-arquivos07\\dirgerti\\SEDT_AUTORE\\TN_AUTORE\\SUSTENTAÇÃO\\Incidentes\\robo_sigs\\TIPO_PRODUTO_ABEND.xlsx', shee_tname='Planilha1')
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



#tipifica os incidentes
def main():
    #EX.main(3)
    driver = inicia_tipificacao()
    print("Fim da execução=======================")
    if (driver != 0):
        driver.quit()
import interacoes_sigs_v7 as sigs
import datetime

def invoca_robo():
    print('Digite\n 1 para extração de Incidentes\n 2 para triagem\n 3 para tipificação de incidentes\n 4 para horas trabalhadas\n 5 para histórico de incidentes')
    comando = int(input('Digite aqui'))
    if comando ==1:
        sigs.main_extrai_incidentes(1)#faz a extração completa
    elif comando ==2:#triagem
        sigs.main_triagem()
    elif comando ==3:#tipificação
        sigs.main_tipifica_incidentes()
    elif comando ==4:#horas trabalhadas
        sigs.main_horas_trabalhadas()
    elif comando ==5:#histórico de incidentes
        tipo_historico = int(input('Digite 1 se deseja realizar uma nova extração de incidentes ou 2 se deseja utilizar o arquivo atual da rede'))
        if tipo_historico==1:
            sigs.main_extrai_incidentes(1)
        sigs.gera_historico_incidentes()
    else:
        print('Comando inválido.')
        invoca_robo()
    now = datetime.datetime.now()
    hora = now.strftime("%H:%M:%S")
    print('Fim da execução às ' + hora)

invoca_robo()
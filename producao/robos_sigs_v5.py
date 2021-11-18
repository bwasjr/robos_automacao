import interacoes_sigs_v7 as sigs
import datetime


def invoca_robo():
    print('Digite\n 1 para extração de Incidentes\n 2 para triagem\n 3 para tipificação de incidentes\n 4 para horas trabalhadas\n 5 para histórico de incidentes')
    comando = int(input('Digite aqui'))
    resultado = sigs.main(comando)
    if resultado == 'invalido':
        invoca_robo()
    agora = datetime.datetime.now()
    hora = agora.strftime("%H:%M:%S")
    print('Fim da execução às ' + hora)


invoca_robo()

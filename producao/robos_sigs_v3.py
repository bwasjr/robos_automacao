import interacoes_sigs_v1 as sigs

def invoca_robo():
    comando = int(input('Digite 1 para extração de Incidentes\n 2 para triagem\n 3 para tipificação de incidentes'))
    if comando ==1:
        sigs.main_extrai_incidentes(1)#faz a extração completa
    elif comando ==2:#triagem
        sigs.main_triagem()
    elif comando ==3:#tipificação
        sigs.main_tipifica_incidentes()
    else:
        print('Comando inválido.')
        invoca_robo()

invoca_robo()
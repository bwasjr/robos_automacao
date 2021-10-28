import robo_triagem_v16 as TRI
import robo_sigs_extrai_incidentes_v19 as EXT
import tipifica_incidentes_v4 as TIP

def invoca_robo():
    comando = int(input('Digite 1 para extração de Incidentes\n 2 para triagem\n 3 para tipificação de incidentes'))
    if comando ==1:
        EXT.main(1)#faz a extração completa
    elif comando ==2:#triagem
        TRI.main()
    elif comando ==3:#tipificação
        TIP.main()
    else:
        print('Comando inválido.')
        invoca_robo()

invoca_robo()
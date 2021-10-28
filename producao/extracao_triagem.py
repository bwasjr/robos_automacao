import robo_triagem_v14 as TRI
import robo_sigs_extrai_incidentes_v18 as EXT

def invoca_robo():
    comando = int(input('Digite 1 para extração do SIGS ou 2 para triagem de incidentes'))
    if comando ==1:
        EXT.main(1)#faz a extração completa
    elif comando ==2:
        TRI.main()
    else:
        print('Comando inválido.')
        invoca_robo()

invoca_robo()
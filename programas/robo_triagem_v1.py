import robo_sigs_extrai_ins_sust_bare_v1 as r1

r1.main()#extrai os incidentes do grupo sustentacao bare

def loga_pag_pesquisa():#loga e acessa a página de pesquisa
    driver = r1.instancia_driver()
    driver.get('http://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')#acessa a pagina do SIGS
    r1.login(driver)#executa o login
    r1.painel_esquerda(driver)

loga_pag_pesquisa()#loga no SIGS e navega até a área de pesquisar incidentes

#abre o arquivo e começa a ler a coluna dos incidentes
    #le o incidente da planilha
    #digita o numero do incidente na planilha e pesquisa
    #le a descrição do incidente
    #procura as palavras chave da tabela de deXpara na descrição do incidente
    #se palavra encontrada então
        #redireciona para o grupo adequado
        #retorna para a pesquisa de incidentes
    #senão, retorna para a pesquisa de incidentes

print("=======================\nFim da triagem\n=======================")
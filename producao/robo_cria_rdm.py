import interacoes_sigs_v8alpha as SIGS
import interacoes_selenium as IS

def loga_sigs():
    driver = SIGS.instancia_driver()
    driver.get('https://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')#acessa a pagina do SIGS
    SIGS.login(driver)#executa o login
    return driver

def main():
    driver = loga_sigs()
    SIGS.painel_esquerda(driver, 3, ['Gerenciamento de Mudanças', 'Mudanças', 'Pesquisar Mudanças'], 4)
    SIGS.time.sleep(4)
    IS.clica_xpath_time(driver, "//*[contains(text(), 'Usar Padrão')]", 5)
    SIGS.time.sleep(4)
    #lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')#obtem a lista dos iframes
    #IS.troca_frame(driver, lista_objetos[-2])#seleciona o iframe do formulario de pesquisa do incidente
    lista_objetos = IS.retorna_objetos(driver, 'tag', 'iframe')
    for objeto in lista_objetos:
        print(objeto.get_attribute('id'))
    driver.switch_to.frame(lista_objetos[1])
    IS.insere_texto_xpath(driver, "//*[@id='X11']",'653986')
    driver.switch_to.default_content()
    IS.clica_xpath(driver, '//button[text()="Pesquisar"]')
    SIGS.time.sleep(4)
    html = driver.execute_script("return document.getElementsByTagName('html')[0].innerHTML")
    
    f = open("source_pagina_html.txt", "a", encoding="utf-8")
    f.write(html)
    f.close()
    
    print(html)
    SIGS.time.sleep(300)
    driver.quit()
    print('fim de execução')

main()

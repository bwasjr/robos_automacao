from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


def instancia_driver():  # INICIALIZACAO do chromedriver
    options = webdriver.ChromeOptions()
    #options.headless = True
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


def clica_id(driver, id):
    element = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, id)))
    element.click()


def clica_id_time(driver, id, timeout):
    element = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.ID, id)))
    element.click()


def clica_time(driver, id, timeout):
    element = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.ID, id)))
    element.click()


def clica_xpath(driver, xpath):
    element = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()


def clica_xpath_time(driver, xpath, timeout):
    element = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()


def clica_objeto(driver, objeto):
    element = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, objeto.get_attribute("id"))))
    element.click()


def clica_classe(driver, classe, posicao):
    lista = driver.find_elements_by_class_name(classe)
    lista[posicao].click()  # clica no elemento desejado


def clica_objeto_lista(lista, posicao_objeto):
    lista[posicao_objeto].click()


def clica_por_texto(lista_objetos, texto):
    limite = 0  # variavel de controle de timout
    clicado = False  # variavel de controle que determina que o objeto foi clicado
    # loop para esperar o objeto ficar disponível para clicar
    while (limite < 21 and clicado == False):
        for botao in lista_objetos:
            if ((botao.text == texto) and (botao.is_enabled and botao.is_displayed)):
                botao.click()
                clicado = True
        time.sleep(1)
        limite += 1


def clica_por_texto_time(lista_objetos, texto, timeout):
    limite = 0  # variavel de controle de timout
    clicado = False  # variavel de controle que determina que o objeto foi clicado
    # loop para esperar o objeto ficar disponível para clicar
    while (limite < timeout and clicado == False):
        for botao in lista_objetos:
            if ((botao.text == texto) and (botao.is_enabled and botao.is_displayed)):
                botao.click()
                clicado = True
        time.sleep(1)
        limite += 1


def insere_texto(elemento, texto):
    elemento.send_keys(texto)


def insere_texto_xpath(driver, xpath, texto):
    elemento = retorna_objetos(driver, 'xpath', xpath)
    elemento[0].send_keys(texto)


def troca_frame(driver, frame):
    driver.switch_to.frame(frame)


def retorna_objetos(driver, selecao, nome):
    if (selecao == 'id'):
        return driver.find_element_by_id(nome)
    elif (selecao == 'class'):
        return driver.find_elements_by_class_name(nome)
    elif (selecao == 'name'):
        return driver.find_element_by_name(nome)
    elif ((selecao == 'tag')):
        return driver.find_elements_by_tag_name(nome)
    elif ((selecao == 'xpath')):
        return driver.find_elements_by_xpath(nome)

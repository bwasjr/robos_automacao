from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def instancia_driver():    #INICIALIZACAO
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument("--test-type")
    options.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": "\\\\srv-arquivos07\dirgerti\SEDT_AUTORE\TN_AUTORE\SUSTENTAÇÃO\Incidentes\Dashboard Incidentes",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    options.add_experimental_option('prefs', prefs)

    driver = webdriver.Chrome('C:\\Users\\g571602\\Documents\\Python\\selenium\\chromedriver.exe', options=options)
    return driver

def clica_id(driver, id):
    element = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.ID, id)))
    element.click()

def clica_ok(driver, id):
    element = WebDriverWait(driver, 3).until(
    EC.element_to_be_clickable((By.ID, id)))
    element.click()

def clica_xpath(driver, xpath):
    element = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()

def clica_objeto(driver, objeto):
    element = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.ID, objeto.get_attribute("id"))))
    element.click()

def clica_classe(driver, classe, posicao):
    lista = driver.find_elements_by_class_name(classe)
    lista[posicao].click() #clica no elemento desejado

def clica_objeto_lista(lista, posicao_objeto):
    lista[posicao_objeto].click()

def clica_por_texto(lista_objetos,texto):
    limite = 0 #variavel de controle de timout
    clicado = False #variavel de controle que determina que o objeto foi clicado
    while (limite <21 and clicado == False):#loop para esperar o objeto ficar disponível para clicar
        for botao in lista_objetos:
            if ((botao.text == texto) and (botao.is_enabled and botao.is_displayed)):
                botao.click()
                clicado = True
        time.sleep(1)
        limite +=1

def insere_texto(elemento, texto):
    elemento.send_keys(texto)

def troca_frame(driver, frame):
    driver.switch_to.frame(frame)

def retorna_objetos(driver, selecao, nome):
    if (selecao == 'id'):
        return driver.find_element_by_id(nome)
    if (selecao == 'class'):
        return driver.find_elements_by_class_name(nome)
    if (selecao == 'name'):
        return driver.find_element_by_name(nome)
    if ((selecao == 'tag')):
        return driver.find_elements_by_tag_name(nome)
    if ((selecao == 'xpath')):
        return driver.find_elements_by_xpath(nome)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import numpy as np
import time
from datetime import date
from datetime import datetime
from datetime import timedelta
import os
import glob
import pandas as pd
import csv
from xlsxwriter.workbook import Workbook
import interacoes_selenium as IS
import interacoes_sigs as sigs
from sqlalchemy import create_engine

# import pymysql
import datetime
import sys
from dateutil.relativedelta import *

driver = sigs.instancia_driver()
# acessa a pagina do SIGS
driver.get('https://servicemanager.net.bradesco.com.br/SM/index.do?lang=pt-Br')
sigs.login(driver)  # executa o login

try:
    IS.clica_xpath_time(driver, '//button[text()="Não"]', 3)
    # print('clicou no botão não')
except TimeoutException:
    print('A janela de prompt para salvar não surgiu, o robô pode continuar')
    pass

import shutil
import sys
from datetime import datetime
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
from pathlib import Path
import os

dataset = pd.read_excel('input.xlsx')
print(dataset.dtypes)

url = 'http://gestaodepessoal.santamaria.rs.gov.br/portalservidor#/'
url_ipassp = 'http://www.ipasspsm.net/PortalServidor/#/'

# ITERAÇÃO POR MATRÍCULA
for linha in dataset.iterrows():
    # CONFIGURA O NAVEGADOR
    navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    navegador.implicitly_wait(120)
    actions = ActionChains(navegador)

    # ATRIBUTOS
    print('*' * 20)
    print(linha[1])
    nome = linha[1]['Nome']
    matricula = str(linha[1]['Matrícula'])
    login = linha[1]['Login']
    senha = linha[1]['Senha']
    termo_inicial = int(linha[1]['Termo Inicial'])
    portal = linha[1]['PORTAL']
    caminho_cliente = linha[1]['caminho']


    mes_atual = '12'


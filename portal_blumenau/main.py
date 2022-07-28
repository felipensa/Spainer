import shutil
from datetime import datetime
from time import sleep
import pyautogui as pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from pathlib import Path


# ARMAZENA OS DADOS
dataset = pd.read_excel('Parametros_BluBot.xls')
dataset['Ano inicial'] = dataset['Ano inicial'].astype(int)
print(dataset.dtypes)

# CONFIGURA O NAVEGADOR
navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
navegador.implicitly_wait(10)
actions = ActionChains(navegador)

# ITERAÇÃO POR MATRÍCULA
for linha in dataset.iterrows():
    # ATRIBUTOS
    print('*' * 20)
    print(linha[1])

    matricula = str(linha[1]['Matrícula'])
    login = linha[1]['Login']
    senha = linha[1]['Senha']
    ano_inicial = linha[1]['Ano inicial']
    ano_final = linha[1]['Ano final']
    caminho = linha[1]['Caminho para pasta do cliente']
    periodo = ano_final - ano_inicial + 1
    mes_atual = '12'

    # ACESSO PÁGINA
    link = 'https://senior.blumenau.sc.gov.br/restrito/'
    navegador.get(url=link)

    # ACESSO LOGIN
    navegador.find_element(by=By.XPATH, value='//*[@id="txtNomUsu"]').send_keys(login)
    navegador.find_element(by=By.XPATH, value='//*[@id="txtSenUsu"]').send_keys(senha)
    navegador.find_element(by=By.XPATH, value='//*[@id="submit"]').click()
    sleep(3)

    # FECHA O NOVO NAVEGADOR
    pyautogui.keyDown('ctrl')
    pyautogui.press('w')
    pyautogui.keyUp('ctrl')
    sleep(1)

    # TROCA GUIA PARA PDF
    pyautogui.keyDown('ctrl')
    pyautogui.press('tab')
    pyautogui.keyUp('ctrl')

    # ACESSO E DOWNLOAD FICHA FINANCEIRA
    for ano in range(0, periodo):
        if ano_inicial == '2022':
            data = datetime.datetime.now()
            mes_atual = int(data.strftime("%m"))
            mes_atual -= 1
            if mes_atual < 10:
                mes_atual = '0' + str(mes_atual)
                str(mes_atual)
            else:
                str(mes_atual)

        print(mes_atual)
        print(ano_inicial)
        ano_inicial = str(ano_inicial)

        link_financeiro = 'https://senior.blumenau.sc.gov.br/restrito/conector?ACAO=EXEREL&SIS=FP&NOME=FPFF512' \
                          '.OPE&dado_EMosUsu=S&dado_ENTipCal=3&dado_ECalcMed=S&dado_EspNivTot=0&dado_EspNivQue=0&' \
                          'dado_ENSomTot=N&dado_EAbrEmp=' + login[0] + '&dado_EAbrTcl=' + login[2] + '&dado_EAb' \
                          'rCad=' + str(matricula) + '&dado_EAbrTco=1-99&dado_EAbrTsa=1-99&dado_EAbrSit=1-999&dado_EAbr' \
                          'Eve=1-9999&dado_EAbrNEv=0&LINWEB=&dado_EListarRef=S&dado_ENPerIni=01/' + str(ano_inicial) +\
                          '&dado_EDatRef=' + mes_atual + '/' + str(ano_inicial)
        navegador.get(url=link_financeiro)
        sleep(3)

        # SALVA ARQUIVO
        pyautogui.keyDown('ctrl')
        pyautogui.press('s')
        pyautogui.keyUp('ctrl')
        sleep(2)

        # NOMEIA ARQUIVOS E MOVE PARA PASTA DE FICHAS
        pyautogui.write('ficha financeira ' + matricula + ' ' + ano_inicial)
        pyautogui.press(['enter'])
        sleep(1)
        downloads = caminho = Path.home() / "Downloads"

        shutil.move(downloads / f"ficha financeira {matricula} {ano_inicial}.pdf",
                    './Fichas ' + matricula)

        ano_inicial = int(ano_inicial)
        ano_inicial += 1
        ano_inicial = str(ano_inicial)

    # UNIFICA OS PDFS
    from junta_pdf import juntaPDF
    juntaPDF(matricula)
    
    # CONVERTE O PDF CONSOLIDADO EM EXCEL
    from adobeSimples import conversao_excel
    conversao_excel(navegador, matricula, downloads)

    # FECHA O NAVEGADOR
    navegador.quit()

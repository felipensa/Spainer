import shutil
import sys
from datetime import datetime
from time import sleep
import pyautogui as pyautogui
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

sys.tracebacklimit = 0

# ARMAZENA OS DADOS
dataset = pd.read_excel('Parametros_BluBot.xlsx')
dataset['Ano inicial'] = dataset['Ano inicial'].astype(int)
dataset['Matrícula'] = dataset['Matrícula'].astype(str)
dataset['Ano final'] = dataset['Ano final'].astype(int)

print(dataset.dtypes)

print('Iniciando iteraçao...')
ponto_inicial = Path.cwd()

# ITERAÇÃO POR MATRÍCULA
for linha in dataset.iterrows():
    # CONFIGURA O NAVEGADOR
    navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    navegador.implicitly_wait(120)
    actions = ActionChains(navegador)

    # ATRIBUTOS
    print('*' * 20)
    print(linha[1])

    matricula = str(linha[1]['Matrícula'])
    login = linha[1]['Login']
    senha = linha[1]['Senha']
    ano_inicial = int(linha[1]['Ano inicial'])
    ano_final = int(linha[1]['Ano final'])
    nome = linha[1]['Nome']
    periodo = ano_final - ano_inicial + 1
    mes_atual = '12'

    # CRIA PASTA DE FICHAS
    pasta_fichas = rf'P:\NAVARRO\CLIENTES\{nome}\HORA ATIVIDADE\Fichas {matricula}'
    Path.mkdir(Path(pasta_fichas), parents=True)
    pasta_cumprimento = rf'P:\NAVARRO\CLIENTES\{nome}\HORA ATIVIDADE\Cumprimento de sentença'
    Path.mkdir(Path(pasta_cumprimento))

    # ACESSO PÁGINA
    link = 'https://senior.blumenau.sc.gov.br/restrito/login.htm'
    navegador.get(url=link)

    # ACESSO LOGIN
    navegador.find_element(by=By.XPATH, value='//*[@id="txtNomUsu"]').send_keys(login)
    navegador.find_element(by=By.XPATH, value='//*[@id="txtSenUsu"]').send_keys(senha)
    navegador.find_element(by=By.XPATH, value='//*[@id="submit"]').click()
    navegador.switch_to.window(navegador.window_handles[1])

    consulta_janelas = len(navegador.window_handles)
    while consulta_janelas < 3:
        try:
            print('entrou no try...')
            WebDriverWait(navegador, 20).until(EC.alert_is_present())
            alert = navegador.switch_to.alert
            alert_text = alert.text
            print(alert_text)
            alert.accept()
            print("Falha no login")
            falha = True
            break
        except TimeoutException:
            print("Possível login...")
            consulta_janelas = len(navegador.window_handles)
            falha = False

    if falha:
        continue

    sleep(5)
    navegador.window_handles[2]
    # FECHA O NOVO NAVEGADOR
    navegador.close()

    # TROCA GUIA PARA PDF
    navegador.switch_to.window(navegador.window_handles[0])

    # ACESSO E DOWNLOAD FICHA FINANCEIRA
    for ano in range(0, periodo):
        if ano_inicial == '2022':
            data = datetime.now()
            mes_atual = int(data.strftime("%m"))
            mes_atual -= 1
            if mes_atual < 10:
                mes_atual = '0' + str(mes_atual)
                str(mes_atual)
            else:
                str(mes_atual)

        print(f'Mês atual: {mes_atual} | Ano inicial {ano_inicial}')
        ano_inicial = str(ano_inicial)

        link_financeiro = 'https://senior.blumenau.sc.gov.br/restrito/conector?ACAO=EXEREL&SIS=FP&NOME=FPFF512' \
                          '.OPE&dado_EMosUsu=S&dado_ENTipCal=3&dado_ECalcMed=S&dado_EspNivTot=0&dado_EspNivQue=0&' \
                          'dado_ENSomTot=N&dado_EAbrEmp=' + login[0] + '&dado_EAbrTcl=' + login[2] + '&dado_EAb' \
                                                                                                     'rCad=' + \
                          str(matricula) + '&dado_EAbrTco=1-99&dado_EAbrTsa=1-99&dado_EAbrSit=1-999&dado_EAbr' \
                          'Eve=1-9999&dado_EAbrNEv=0&LINWEB=&dado_EListarRef=S&dado_ENPerIni=01/' + str(ano_inicial) + \
                          '&dado_EDatRef=' + mes_atual + '/' + str(ano_inicial)
        navegador.get(url=link_financeiro)

        try:
            print('Confere período...')
            WebDriverWait(navegador, 20).until(EC.alert_is_present())
            alert = navegador.switch_to.alert
            alert_text = alert.text
            print(alert_text)
            alert.accept()
            print("Provável aposentadoria ou exoneração")
            falha_periodo = True
        except TimeoutException:
            print("Contracheque disponível!")
            falha_periodo = False

        if falha_periodo:
            continue

        sleep(3)

        # SALVA ARQUIVO
        pyautogui.keyDown('ctrl')
        pyautogui.press('s')
        pyautogui.keyUp('ctrl')
        sleep(2)

        # NOMEIA ARQUIVOS E MOVE PARA PASTA DE FICHAS
        pyautogui.write('ficha financeira ' + matricula + ' ' + ano_inicial + '.pdf')
        pyautogui.press(['enter'])
        sleep(1)
        downloads = Path.home() / "Downloads"

        shutil.move(str(downloads) + fr"\ficha financeira {matricula} {ano_inicial}.pdf", pasta_fichas)
        shutil.copyfile(str(pasta_fichas) + fr"\ficha financeira {matricula} {ano_inicial}.pdf",
                        str(pasta_cumprimento) + rf'\5 Fichas financeiras.pdf')

        ano_inicial = int(ano_inicial)
        ano_inicial += 1
        ano_inicial = str(ano_inicial)

    # UNIFICA OS PDFS
    from junta_pdf import juntaPDF

    juntaPDF(matricula, ponto_inicial, pasta_fichas)

    # CONVERTE O PDF CONSOLIDADO EM EXCEL
    from adobeSimples import conversao_excel

    conversao_excel(navegador, matricula, downloads, ponto_inicial, pasta_fichas)

# FECHA O NAVEGADOR
navegador.quit()
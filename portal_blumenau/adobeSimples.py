import sys
from pathlib import Path
from time import sleep
import pyautogui
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


def conversao_excel(navegador, matricula, pasta_downloads, pasta_inicial):
    # ACESSA A PÁGINA DO ADOBE
    link = 'https:/www.adobe.com/br/acrobat/online/pdf-to-excel.html'
    navegador.get(url=link)
    sleep(5)

    # PEDE O ARQUIVO DE INPUT
    navegador.find_element(by=By.XPATH, value='//*[@id="lifecycle-nativebutton"]').click()
    sleep(5)

    # INFORMA O ARQUIVO E ORDENA CONVERSÃO
    caminho = str(pasta_inicial) + r'\Fichas ' + matricula
    pyautogui.write(caminho)
    pyautogui.press('Enter')
    sleep(2)
    pyautogui.write('fichas_financeiras ' + matricula + '.pdf')
    pyautogui.press('Enter')
    sleep(1)

    # ESPERA CLASS APARECER
    wait_for_element = 60  # ESPERA TIMEOUT EM SEGUNDOS
    try:
        WebDriverWait(navegador, wait_for_element).until(
            EC.element_to_be_clickable((By.CLASS_NAME,
                                        "spectrum-Button spectrum-Button--cta "
                                        "DownloadOrShare__downloadButton___3z1LR")))
    except TimeoutException as e:
        print("Wait Timed out")
        print(e)

    # DOWNLOAD
    navegador.find_element(by=By.XPATH, value='//*[@id="dc-hosted-ec386752"]/div/'
                                              'div/div[2]/div/section[2]/div/div[1]/div[2]/button[1]').click()
    sleep(5)

    # MOVE ARQUIVO DE DOWNLOADS PARA PASTA ADEQUADA
    shutil.move(str(pasta_downloads) + fr"\fichas_financeiras {matricula}.xlsx",
                str(pasta_inicial) + fr"\Fichas {matricula}" + fr'\fichas_financeiras {matricula}.xlsx')

    navegador.quit()

    return 1

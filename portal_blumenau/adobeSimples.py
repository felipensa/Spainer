import sys
from pathlib import Path
from time import sleep
import pyautogui
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


def conversao_excel(navegador, matricula, pasta_downloads, pasta_fichas):
    print('Convertendo fichas financeiras consolidadas...')
    # ACESSA A PÁGINA DO ADOBE
    link = 'https:/www.adobe.com/br/acrobat/online/pdf-to-excel.html'
    navegador.get(url=link)
    sleep(5)

    # INPUT DO ARQUIVO CONSOLIDADO
    navegador.find_element(by=By.XPATH, value='//*[@id="fileInput"]').send_keys(str(pasta_fichas)
                                                                                + rf'\fichas_financeiras {matricula}.pdf')

    sleep(5)

    # ESPERA CLASS APARECER
    wait_for_element = 60  # ESPERA TIMEOUT EM SEGUNDOS
    try:
        WebDriverWait(navegador, wait_for_element).until(
            EC.element_to_be_clickable((By.XPATH,
                                        '//*[@id="dc-hosted-ec386752"]/div/div/div[2]/div'
                                        '/section[1]/div/div/div[2]/div[1]/button[1]')))
        print('Botão de download clicável...')

    except TimeoutException as e:
        print("Tempo de espera expirado...")

    # DOWNLOAD
    navegador.find_element(by=By.XPATH, value='//*[@id="dc-hosted-ec386752"]/div/div/div[2]/div'
                                              '/section[1]/div/div/div[2]/div[1]/button[1]').click()
    sleep(5)

    # MOVE ARQUIVO DE DOWNLOADS PARA PASTA ADEQUADA
    shutil.move(pasta_downloads + fr"\fichas_financeiras {matricula}.xlsx",
                str(pasta_fichas) + fr'\fichas_financeiras {matricula}.xlsx')

    print('Conversão finalizada!')
    navegador.quit()

    return 1

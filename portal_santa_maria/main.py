import shutil
from time import sleep
import datetime as datetime
import pyautogui as pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
from pathlib import Path
from PyPDF2 import PdfFileMerger
import os


def clica_xpath(driver, xpath):
    driver.find_element(by=By.XPATH, value=xpath).click()


def aguarda_download():
    while not Path.is_file(downloads / "contracheque.pdf"):
        pass


def junta_pdf(matricula, nome, ponto_inicial):
    # GUARDA O CAMINHO E MUDA O AMBIENTE DO PYTHON PARA TAL PASTA
    pasta = fr'./Fichas {matricula} - {nome}'
    os.chdir(pasta)

    # ARMAZENA OS PDFS INDIVIDUALMENTE
    x = [a for a in os.listdir() if a.endswith(".pdf")]

    # ORDENA ARQUIVOS
    def mergeSort(alist):
        if len(alist) > 1:
            mid = len(alist) // 2
            lefthalf = alist[:mid]
            righthalf = alist[mid:]

            mergeSort(lefthalf)
            mergeSort(righthalf)

            i = 0
            j = 0
            k = 0

            while i < len(lefthalf) and j < len(righthalf):
                if lefthalf[i] < righthalf[j]:
                    alist[k] = lefthalf[i]
                    i = i + 1
                else:
                    alist[k] = righthalf[j]
                    j = j + 1
                k = k + 1

            while i < len(lefthalf):
                alist[k] = lefthalf[i]
                i = i + 1
                k = k + 1

            while j < len(righthalf):
                alist[k] = righthalf[j]
                j = j + 1
                k = k + 1
        print("Merging ", x)

    mergeSort(x)
    print(x)

    # UNIFICA OS ARQUIVOS
    merger = PdfFileMerger()
    for pdf in x:
        merger.append(open(pdf, 'rb'))

    # SALVA O ARQUIVO CONSOLIDADO
    with open(f"fichas_financeiras {matricula}.pdf", "wb") as fout:
        merger.write(fout)

    os.chdir(str(ponto_inicial))
    return print('Arquivos unificados...')


def conversao_excel(navegador, matricula, pasta_downloads, pasta_inicial, nome):
    # ACESSA A PÁGINA DO ADOBE
    link_adobe = 'https:/www.adobe.com/br/acrobat/online/pdf-to-excel.html'
    navegador.get(url=link_adobe)

    # PEDE O ARQUIVO DE INPUT
    navegador.find_element(by=By.XPATH, value='//*[@id="lifecycle-nativebutton"]').click()
    sleep(5)

    # INFORMA O ARQUIVO E ORDENA CONVERSÃO
    pdf_consolidado = str(pasta_inicial) + rf'\Fichas {matricula} - {nome}'
    pyautogui.write(pdf_consolidado)
    pyautogui.press('Enter')
    sleep(2)
    pyautogui.write(f'fichas_financeiras {matricula}.pdf')
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

    # DOWNLOAD
    clica_xpath(navegador, '//*[@id="dc-hosted-ec386752"]/div/'
                           'div/div[2]/div/section[2]/div/div[1]/div[2]/button[1]')
    sleep(5)

    # MOVE ARQUIVO DE DOWNLOADS PARA PASTA ADEQUADA
    shutil.move(str(pasta_downloads) + fr"\fichas_financeiras {matricula}.xlsx",
                str(pasta_inicial) + fr"\Fichas {matricula} - {nome}" + fr'\fichas_financeiras {matricula}.xlsx')

    navegador.quit()

    return print('Arquivos convertidos em excel...')


# SALVA PASTA DE TRABALHO
ponto_inicial = Path.cwd()
downloads = Path.home() / "Downloads"

# ARMAZENA OS DADOS
dataset = pd.read_excel('input.xlsx')
dataset['matricula'] = dataset['matricula'].astype(str)

print(dataset.dtypes)

cont_linha = 1
# ITERAÇÃO POR MATRÍCULA
for linha in dataset.iterrows():
    # CONFIGURA O NAVEGADOR
    navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    navegador.implicitly_wait(120)
    actions = ActionChains(navegador)

    # ATRIBUTOS
    print('*' * 20)
    print(linha[1])
    nome = linha[1]['nome']
    matricula = str(linha[1]['matricula'])
    login = linha[1]['login']
    senha = linha[1]['senha']
    mes_inicial = int(linha[1]['mes_inicial'])
    ano_inicial = int(linha[1]['ano_inicial'])
    caminho = linha[1]['caminho']

    # CONTAGEM
    # TESTE CALCULO DE DATAS
    from datetime import datetime
    from dateutil import relativedelta
    from datetime import datetime

    mes_atual = datetime.now().strftime("%m")
    ano_atual = datetime.now().strftime("%Y")
    print((mes_atual, ano_atual))

    # get two dates
    d1 = f'01/{mes_inicial}/{ano_inicial}'
    d2 = f'01/{mes_atual}/{ano_atual}'

    # convert string to date object
    start_date = datetime.strptime(str(d1), "%d/%m/%Y")
    end_date = datetime.strptime(str(d2), "%d/%m/%Y")

    # Get the relativedelta between two dates
    delta = relativedelta.relativedelta(end_date, start_date)

    # get months difference
    res_months = delta.months + (delta.years * 12)
    print('Total Months between two dates is:', res_months)

    # CÁLCULO: QUANTIDADE DE MESES
    periodo = (12 - mes_inicial + 1) + ((int(ano_atual) - ano_inicial - 1) * 12) + (int(mes_atual) - 1)
    print(f'Período: {periodo} meses')

    # ACESSO PÁGINA
    link = 'http://gestaodepessoal.santamaria.rs.gov.br/portalservidor#/'
    navegador.get(url=link)

    # ACESSO LOGIN
    navegador.find_element(by=By.XPATH, value='//*[@id="content"]/form/fieldset/div[1]/input').send_keys(login)
    navegador.find_element(by=By.XPATH, value='//*[@id="content"]/form/fieldset/div[2]/div/input').send_keys(senha)
    clica_xpath(navegador, '//*[@id="content"]/form/fieldset/input[3]')  # CLICA EM ENTRAR

    # TESTA VALIDADE DAS CREDENCIAIS
    navegador.implicitly_wait(2)
    try:
        print(navegador.find_element(by=By.XPATH, value='//*[@id="toasty-container"]/div/div/span[1]').text)
        continue
    except NoSuchElementException:
        pass
        navegador.implicitly_wait(120)
        clica_xpath(navegador, '//*[@id="options"]/ul/li[1]/a')

    # CRIA PASTA DE FICHAS
    os.mkdir(f'Fichas {matricula} - {nome}')

    # ITERAÇÃO POR MESES
    cont = 0
    for mes in range(cont, periodo):
        clica_xpath(navegador,
                    '//*[@id="content"]/div[3]/form/fieldset/div/div[1]/div[1]/span/span/span[2]/span')  # ABRE LISTA ANO
        sleep(0.8)
        primeiro_valor = int(navegador.find_element(by=By.XPATH, value='//*[@id="ano_listbox"]/li[2]').text)
        quantidade_anos = ano_inicial - int(primeiro_valor)

        clica_xpath(navegador, f'//*[@id="ano_listbox"]/li[{quantidade_anos + 2}]')  # SELECIONA ELEMENTO
        clica_xpath(navegador,
                    '//*[@id="content"]/div[3]/form/fieldset/div/div[2]/div/span/span/span[2]/span')  # ABRE LISTA MES
        sleep(0.8)
        clica_xpath(navegador, f'//*[@id="mes_listbox"]/li[{mes_inicial + 1}]')  # SELECIONA ELEMENTO
        clica_xpath(navegador, '//*[@id="content"]/div[3]/form/button[1]')  # PESQUISAR
        sleep(0.8)

        # TRATAMENTO PARA 13º
        if mes_inicial != 12:
            clica_xpath(navegador, '//*[@id="content"]/div[3]/div/table/tbody/tr/td[1]/a[1]')  # BAIXA DOCUMENTO
            aguarda_download()
            shutil.move(downloads / "contracheque.pdf", str(ponto_inicial) +
                        rf'\Fichas {matricula} - {nome}\Ficha {ano_inicial} - {mes_inicial}.pdf')

        else:
            clica_xpath(navegador, '//*[@id="content"]/div[3]/div/table/tbody/tr/td[1]/a[1]')  # BAIXA DOCUMENTO 13º
            aguarda_download()
            shutil.move(downloads / "contracheque.pdf", str(ponto_inicial) +
                        rf'\Fichas {matricula} - {nome}\Ficha {ano_inicial} - {mes_inicial + 1}.pdf')
            sleep(5)
            clica_xpath(navegador, '//*[@id="content"]/div[3]/div/table/tbody/tr[2]/td[1]/a[1]')
            aguarda_download()
            shutil.move(downloads / "contracheque.pdf", str(ponto_inicial) +
                        rf'\Fichas {matricula} - {nome}\Ficha {ano_inicial} - {mes_inicial}.pdf')

        # CONTADORES
        print(f'Mês: {mes_inicial} / Ano: {ano_inicial}')
        print(f'Contagem meses: {mes} / {periodo}')

        # INCREMENTO
        if mes_inicial == 12:
            mes_inicial = 0
            ano_inicial += 1
        mes_inicial += 1
        sleep(5)

    # CONTADORES GERAIS
    print(f'Contagem linhas: {cont_linha}/{dataset.shape[0]}')
    cont_linha += 1

    # JUNTA PDF
    junta_pdf(matricula, nome, ponto_inicial)

    # CONVERTE PDF CONSOLIDADO EM EXCEL
    conversao_excel(navegador, matricula, downloads, ponto_inicial, nome)

print('Finalizado!')

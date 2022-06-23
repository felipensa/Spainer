from datetime import datetime
from easygui import boolbox
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium.webdriver.common.by import By
from time import sleep
from selenium import webdriver
import pandas as pd
import ctypes
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook

notas = pd.read_excel('notas.xlsx')
notas.columns = ['notas']
notas.dropna(inplace=True)
notas.reset_index(drop=True, inplace=True)
selecao = []

for linha in notas['notas']:
    if len(linha) <= 60:
        selecao.append(linha)

cont = 0
selecao_final = []
for item in selecao:
    print(item)
    if 'CNJ:' in item and len(item) > 50:
        selecao[cont] = item.split(' ')[3]
        selecao[cont] = selecao[cont].split(')')[0]
        selecao_final.append(selecao[cont])
        print(f'Primeiro: {item}')
    elif 'CNJ' in item and len(item) <= 30:
        selecao[cont] = item.split('(')[0]
        selecao_final.append(selecao[cont])
        print(f'Segundo: {item}')
    elif 'CNJ' in item and len(item) > 30:
        selecao[cont] = item.split(' ')[2]
        selecao[cont] = selecao[cont].split(')')[0]
        selecao_final.append(selecao[cont])
        print(f'Terceiro: {item}')
    else:
        print('Caso fora do PPE')
    cont += 1

cont = 0
processo_sem_format = ''
for processo in selecao_final:
    while len(selecao_final[cont]) < 25:
        selecao_final[cont] = '0' + selecao_final[cont]
    processo_formatado = selecao_final[cont]
    selecao_final[cont] = ''.join(selecao_final[cont].split('.'))
    selecao_final[cont] = ''.join(selecao_final[cont].split('-'))
    selecao_final[cont] = (selecao_final[cont], processo_formatado)
    cont += 1
print(f'Selecação final : {len(selecao_final)}')

dados_preparados = pd.DataFrame(selecao_final, index=None, columns=['sem_formatacao', 'formatado'])
dados_preparados.to_excel('Resultado.xlsx', engine='openpyxl')

""" SCRAPING """

# DADOS
login = '01401116035'
senha = 'sumerios0394'

# INPUT DE DADOS
total_linhas = len(dados_preparados.index)
print(dados_preparados)
print(f'linhas: {total_linhas}')

#  NAVEGADOR
navegador = webdriver.Chrome(ChromeDriverManager().install())
navegador.implicitly_wait(15)
url = 'https://ppe.tjrs.jus.br/ppe/signin'

# LOGIN
navegador.get(url)  # ABRE NAVEGADOR
sleep(2)
navegador.find_element(by=By.XPATH, value='//*[@id="avisos-signin"]/div/div[3]/p-footer/button/span[2]').click()
sleep(1)
navegador.find_element(by=By.XPATH, value='//*[@id="usuario"]').send_keys(login)  # CAMPO LOGIN
sleep(0.5)
navegador.find_element(by=By.XPATH, value='//*[@id="password"]').send_keys(senha)  # CAMPO SENHA
sleep(0.5)
navegador.find_element(by=By.XPATH, value=
'//*[@id="container-principal"]/app-signin/div[3]/div[2]/div/form/button').click()  # BOTAO ENTRAR
sleep(1.5)
try:
    navegador.find_element(by=By.XPATH, value='//*[@id="navBar"]/button/span').click()
except NoSuchElementException:
    pass
sleep(0.5)
navegador.find_element(by=By.XPATH, value='//*[@id="supportedContentDropdownProcesso"]/span').click()
sleep(0.5)
navegador.find_element(by=By.XPATH, value='//*[@id="processo"]/div/a/span').click()
sleep(2)

# LISTAS
cj_num_antigo = []
cj_num_originario = []
cj_processo = []
cj_processoFormatado = []

# PESQUISA PROCESSOS PARA CADA LINHA DO EXCEL
for i in range(0, total_linhas):
    processo = dados_preparados['sem_formatacao'].iloc[i]
    processoFormatado = dados_preparados['formatado'].iloc[i]
    processo = str(processo)
    print(processo)
    print(f'etapa {i + 1}/{total_linhas}')
    sleep(1)

    navegador.find_element(by=By.XPATH, value='//*[@id="numeroProcesso"]').send_keys(processo)  # CAMPO PESQUISA
    sleep(1.5)
    navegador.find_element(by=By.XPATH, value='//*[@id="bt_pesquisar"]').click()  # BOTÃO PESQUISAR
    sleep(2)
    tentativas = 1
    while 1:
        try:
            navegador.find_element(by=By.XPATH, value=
            '//*[@id="processos-grid"]/div/div/table/tbody/tr').click()  # ACESSA RESULTADO
            break
        except:
            tentativas += 1
            navegador.find_element(by=By.XPATH, value='//*[@id="numeroProcesso"]').send_keys(
                processo)  # ENVIA NUMERO PROCESSO PARA CAMPO PESQUISA
            navegador.find_element(by=By.XPATH, value='//*[@id="bt_pesquisar"]').click()  # BOTÃO PESQUISAR
            sleep(2)
    sleep(10)
    print(f"o robô procurou {tentativas} vezes.")
    # COLETA NUMERO ANTIGO E NUMERO ORIGINÁRIO
    try:
        num_antigo = navegador.find_element(by=By.XPATH,
                                            value='//*[@id="campoDetalhesProcesso"]/div[1]/div[2]/span').text
    except NoSuchElementException:
        num_antigo = 'Sem numero antigo'

    try:
        num_originario = navegador.find_element(by=By.XPATH, value=
        '//*[@id="campoDetalhesProcesso"]/div[6]/div[2]/ppe-detalhes-processo-processos-vinculados/div[2]/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr/td[3]/span/td').text
    except NoSuchElementException:
        num_originario = 'Sem dados do processo'

    # APPENDS
    cj_processo.append(processo)  # ARMAZENA NA LISTA O NUMERO DO PROCESSO
    cj_num_antigo.append(num_antigo)  # ARMAZENA NA LISTA O NUMERO ANTIGO
    cj_num_originario.append(num_originario)  # ARMAZENA NA LISTA O NUMERO DO PROCESSO ORIGINÁRIO
    cj_processoFormatado.append(processoFormatado)  # ARMAZENA NA LISTA O NUMERO DO PROCESSO FORMATADO

    sleep(4)
    try:
        navegador.find_element(by=By.XPATH, value='//*[@id="navBar"]/button').click()
    except NoSuchElementException:
        pass
    sleep(0.5)
    navegador.find_element(by=By.XPATH, value='//*[@id="supportedContentDropdownProcesso"]/span').click()
    sleep(0.5)
    navegador.find_element(by=By.XPATH, value='//*[@id="processo"]/div/a/span').click()
    sleep(1)

df_scraping = {'Principal': cj_processo, 'Número antigo': cj_num_antigo, 'Originário': cj_num_originario,
               'Processo Formatado': cj_processoFormatado}  # ARMAZENA OS VALORES EM UMA MATRIZ
scraping = pd.DataFrame(df_scraping,
                        columns=['Principal', 'Antigo', 'Originário', 'Processo Formatado'])  # CRIA O DATAFRAME
scraping.to_excel('./PROCESSOS ATUALIZADOS.xlsx')

navegador.quit()
ctypes.windll.user32.MessageBoxW(0, "Scraping", "Concluído!", 1)

# CONTROLE
print(scraping)

"""# Filtrar e exportar resultado

## Preparação relatório desdobramentos
"""

"""## Atualizar base de dados"""
message = "Atualizar base de dados"
title = "Base de dados"
if boolbox(message, title, ["Sim", "Não"]):
    atualiza = 'sim'
else:
    atualiza = 'nao'

# Tratamento base de dados — desdobramentos
desdobramentos = pd.read_excel('relatorio_desdobramentos.xlsx')
colunas_tabela = list(desdobramentos.columns)
colunas_novas = list(desdobramentos.iloc[2])
dici = {}
for x in range(len(colunas_tabela)):
    dici[colunas_tabela[x]] = colunas_novas[x]
desdobramentos.rename(columns=dici, inplace=True)
desdobramentos.drop(axis=0, index=[0, 1, 2], inplace=True)


def atualiza_base_de_dados():
    pastas_unicas = set(desdobramentos['Pasta do processo'])
    nova_pasta = []
    for unico in pastas_unicas:
        numero = list(desdobramentos['Pasta do processo']).count(unico) + 1
        if numero < 10:
            nova_pasta.append(unico + '.0' + str(numero))
        else:
            nova_pasta.append(unico + '.' + str(numero))
    return nova_pasta


if atualiza == 'sim':
    atualiza_base_de_dados()

resultado_scraping = pd.read_excel('PROCESSOS ATUALIZADOS.xlsx')

"""## Consultar base de dados"""

# Requisitos: Principal não está cadastrado no Espaider; originário está cadastrado no Espaider

consulta_originario_no_principal = []
for proc_format in scraping['Processo Formatado']:
    indice_principal = scraping.index[scraping['Processo Formatado'] == proc_format].tolist()
    originario = scraping.loc[indice_principal[0], 'Originário']
    indice_originario = scraping.index[scraping['Originário'] == originario].tolist()

    if originario in list(desdobramentos['Número CNJ']) and proc_format not in list(desdobramentos['Número CNJ']):
        indice_originario_principal = desdobramentos.index[desdobramentos['Número CNJ'] == originario].tolist()
        consulta_originario_no_principal.append((scraping['Principal'][indice_originario[0]],
                                                 scraping['Originário'][indice_originario[0]],
                                                 desdobramentos.loc[indice_originario_principal[0],
                                                                    'Pasta deste desdobramento'],
                                                 'originario_principal'))
resultado_consulta = pd.DataFrame(consulta_originario_no_principal)
resultado_consulta.to_excel('Resultado consulta.xlsx')
print(f'{len(consulta_originario_no_principal)} processos com número ORIGINÁRIO cadastrados no PRINCIPAL')


"""## Preparar importação"""

# Preparar SmartImport


wb = load_workbook('exemplo.xlsx')
ws = wb.active
for r in dataframe_to_rows(resultado_consulta, index=False, header=False):
    ws.append(r)
    print(r)

wb.save(f"Processos correção PPE - {datetime.today().strftime('%d-%m-%y')}.xlsx")

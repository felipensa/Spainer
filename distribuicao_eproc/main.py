from selenium import webdriver
from selenium.common.exceptions import *
from time import sleep

from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
# import pyautogui
import ctypes

# from easygui import *


def clica_xpath(driver, xpath):
    driver.find_element(by=By.XPATH, value=xpath).click()

# DADOS FIXOS
from webdriver_manager.chrome import ChromeDriverManager

login = 'RS111059'
senha = 'Alegrianoesforço!'

comarca = 'Blumenau'
cnpj_blumenau = '83108357000115'
oab_mauricio = 'RS036798'
oab_thais = 'RS106785'
assunto = 'Sistema Remuneratório e Benefícios'

##########################INPUT DE DADOS########################
msg = "Informes os dados para ajuizamento"
title = "Dados ajuizamento"
fieldNames = ["Valor da ação(com vírgula): ", "CPF cliente(somente números): ", "Qual a pasta dos documentos?: "]
fieldValues = []  # CRIA VETOR PARA RECEBER RESPOSTAS
fieldValues = multenterbox(msg, title, fieldNames)

# GARANTE QUE NENHUM CAMPO FIQUE VAZIO
while 1:
    if fieldValues is None:
        break
    errmsg = ""
    for i in range(len(fieldNames)):
        if fieldValues[i].strip() == "":
            errmsg = errmsg + ('"%s" é um campo obrigatório.\n\n' % fieldNames[i])
    if errmsg == "": break  # no problems found
    fieldValues = multenterbox(errmsg, title, fieldNames, fieldValues)

# RECEBE AS RESPOSTAS
valor_acao = fieldValues[0]
cpf_cliente = fieldValues[1]
caminho = fieldValues[2]

documentos = caminho + r'\*.pdf'

# CAMPOS DE BOOLEAN
message = "Qual o ano executado?"
title = "Ano executado"

message = "Qual a situação funcional?"
title = "Situação funcional"
if boolbox(message, title, ["Ativo/Exonerado", "Inativo"]):
    situacao = 'ativo'
else:
    situacao = 'inativo'

message = "Cliente possui doença grave?"
title = "Tramitação preferencial"
if boolbox(message, title, ["Sim", "Não"]):
    tramit_pref = 'Sim'
else:
    tramit_pref = 'Não'

message = "Cliente é portador de deficiência?"
title = "Tramitação preferencial"
if boolbox(message, title, ["Sim", "Não"]):
    tramit_deficiencia = 'Sim'
else:
    tramit_deficiencia = 'Não'

##########################################################


# LINKS
navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
navegador.implicitly_wait(60)
actions = ActionChains(navegador)
link = 'https://eproc1g.tjsc.jus.br/eproc/'

# ACESSO AO LOGIN
navegador.get(url=link)
sleep(1)

navegador.find_element(by=By.ID, value='txtUsuario').send_keys(login)  # CAMPO LOGIN
navegador.find_element(by=By.ID, value='pwdSenha').send_keys(senha)  # CAMPO SENHA
sleep(0.5)
navegador.find_element(by=By.ID, value='sbmEntrar').click()  # ACESSA PELO LOGIN
sleep(1)

clica_xpath(navegador, '//*[@id="tr0"]')    # SELECIONA PERFIL
sleep(0.5)

clica_xpath(navegador, '//*[@id="main-menu"]/li[11]/a/span[1]')  # DISTRIBUIR INICIAL

Select(navegador.find_element(by=By.ID, value='selLocalidadeJudicial')).select_by_value('2')  # SELECIONA POA

Select(navegador.find_element(by=By.ID, value='selRitoProcesso')).select_by_value('2')  # RITO ORDINÁRIO

Select(navegador.find_element(by=By.ID, value='selIdGrupoCompetencia')).select_by_value('71')  # FAZENDA PÚBLICA

Select(navegador.find_element(by=By.ID, value='selIdClasseJudicial')).select_by_value('0000000029')

Select(navegador.find_element(by=By.ID, value='txtValorCausa')).send_keys(valor_acao)

# INCLUSÃO MAURÍCIO
navegador.Select(navegador.find_element(by=By.ID, value='txtUsuario')).send_keys(oab_mauricio)
sleep(10)

navegador.Select(navegador.find_element(by=By.ID, value='txtUsuario')).send_keys(Keys.ENTER)
sleep(1)

Select(navegador.find_element(by=By.ID, value='btnIncluir')).click()
sleep(1)

# PROSSSEGUIMENTO
clica_xpath(navegador, '//*[@id="btnSalvar"]')  # PROXIMA PAGINA
sleep(1)

navegador.find_element(by=By.XPATH, value='//*[@id="txtFiltroPesquisa"]').send_keys(assunto)  # ASSUNTO
sleep(0.5)

clica_xpath(navegador, '//*[@id="btnPesquisar"]')  # BOTAO PESQUISA ASSUNTO
sleep(4)

clica_xpath(navegador, '//*[@id="011102_anchor"]/span/span')  # SELECIONA O ASSUNTO
sleep(0.5)

clica_xpath(navegador, '//*[@id="btnIncluirAssunto"]')  # INCLUI O ASSUNTO
sleep(1)

clica_xpath(navegador, '//*[@id="btnSalvar"]')  # PRÓXIMA PÁGINA
sleep(1)


##################

navegador.find_element_by_xpath('//*[@id="txtCpfCnpj"]').send_keys(cpf_cliente)  # CAMPO CPF
sleep(0.5)

navegador.find_element_by_xpath('//*[@id="btnConsultarNome"]').click()  # BOTÃO PARA CONSULTAR CPF
sleep(1)

try:
    navegador.find_element_by_xpath('//*[@id="btnIncluir"]').click()  # BOTAO INCLUIR AUTOR
    sleep(1)
except:
    msgbox("Preencha o cadastro do cliente antes de prosseguir")
    navegador.find_element_by_xpath('//*[@id="btnIncluir"]').click()  # BOTAO INCLUIR AUTOR
    sleep(1)

navegador.find_element_by_xpath('//*[@id="btnProxima"]').click()  # PROXIMA PAGINA
sleep(1)

campo_pessoa = Select(navegador.find_element_by_xpath('//*[@id="selTipoPessoa"]'))
campo_pessoa.select_by_value('ENT')
sleep(0.5)

campo_entidade = Select(navegador.find_element_by_xpath('//*[@id="selEntidade"]'))
campo_entidade.select_by_value('771230778800100040000000000395')  # SELECIONA MUNICÍPIO DE BLUMENAU
sleep(0.5)

botao_incluir_reu = navegador.find_element_by_xpath('//*[@id="btnIncluirEnt"]').click()  # BOTAO DE INCLUIR RÉU
sleep(0.5)

if situacao == 'inativo':
    campo_entidade.select_by_value('721474477367875580270239023830')  # SELECIONA ISSBLU
    sleep(0.5)

    campo_principal = Select(navegador.find_element_by_xpath('//*[@id="selPrincipalEnt"]'))
    campo_principal.select_by_value('N')
    sleep(1)

    navegador.find_element_by_xpath('//*[@id="btnIncluirEnt"]').click()  # APERTA BOTAO DE INCLUIR RÉU
    sleep(0.5)

navegador.find_element_by_xpath('//*[@id="btnProxima"]').click()  # PASSA PARA A PROXIMA PAGINA
sleep(2)

if tramit_pref == 'Sim':  # MARCA PREFERÊNCIA POR DOENÇA GRAVE
    navegador.find_element_by_xpath('//*[@id="chkDoencaGrave"]').click()
    sleep(0.5)
if tramit_deficiencia == 'Sim':  # MARCA PREFERÊNCIA POR DEFICIÊNCIA
    navegador.find_element_by_xpath('//*[@id="chkDeficiente"]').click()
    sleep(0.5)

navegador.find_element_by_xpath('//*[@id="_1"]/div/div[2]').click()  # ABRE CAIXA PARA MULTIUPLOAD
sleep(1)

controle_erro = ''


# SELECIONA TODOS OS ARQUIVOS PDF DA PASTA
def erro_upload():
    message = "Erro no upload de alguns arquivos. Resolva-os antes de prosseguir"
    title = "Erro upload"
    if boolbox(message, title, ["Continuar", "Cancelar"]):
        controle_erro = '0'
    else:
        controle_erro = '1'
    return controle_erro


pyautogui.write(caminho)
sleep(0.5)
pyautogui.press('enter')
sleep(2)
pyautogui.press('tab', presses=11, interval=0.5)
sleep(0.5)

pyautogui.keyDown('ctrl')
pyautogui.press('a')
pyautogui.keyUp('ctrl')
pyautogui.press('enter')
sleep(15)


def classificacao_docs():
    # CLASSIFICA TODOS OS ARQUIVOS
    if data_saida == "Sim":
        lista1 = ['PETIÇÃO INICIAL', 'PROCURAÇÃO', 'SUBSTABELECIMENTO', 'IDENTIDADE', 'COMPROVANTE DE RESIDÊNCIA',
                  'OUTROS', 'OUTROS', 'OUTROS', 'OUTROS', 'TÍTULO EXECUTIVO JUDICIAL', 'TÍTULO EXECUTIVO JUDICIAL',
                  'CONTRACHEQUE', 'LAUDO', 'PLANILHA DE CÁLCULO', 'CÁLCULO', 'OUTROS']
        for x in range(1, 17):
            navegador.find_element_by_xpath('//*[@id="txtTipo_' + str(x) + '"]').send_keys(lista1[x - 1])
            sleep(1)
            navegador.find_element_by_xpath('//*[@id="txtTipo_' + str(x) + '"]').send_keys(Keys.ENTER)
            sleep(0.5)
            if x == 6:
                navegador.find_element_by_xpath('//*[@id="txtObservacao_6"]').send_keys('Registro funcional')
            elif x == 7:
                navegador.find_element_by_xpath('//*[@id="txtObservacao_7"]').send_keys('Lei')
            elif x == 8:
                navegador.find_element_by_xpath('//*[@id="txtObservacao_8"]').send_keys('Lei')
            elif x == 9:
                navegador.find_element_by_xpath('//*[@id="txtObservacao_9"]').send_keys('Tabela de reenquadramento')
            elif x == 16:
                navegador.find_element_by_xpath('//*[@id="txtObservacao_16"]').send_keys('Lei')
        if tramit_deficiencia == 'Sim' or tramit_pref == 'Sim':
            navegador.find_element_by_xpath('//*[@id="txtTipo_17"]').send_keys('LAUDO')
            sleep(1)
            navegador.find_element_by_xpath('//*[@id="txtTipo_17"]').send_keys(Keys.ENTER)


    else:
        lista2 = ['PETIÇÃO INICIAL', 'PROCURAÇÃO', 'SUBSTABELECIMENTO', 'IDENTIDADE', 'COMPROVANTE DE RESIDÊNCIA',
                  'OUTROS', 'OUTROS', 'OUTROS', 'OUTROS', 'TÍTULO EXECUTIVO JUDICIAL', 'TÍTULO EXECUTIVO JUDICIAL',
                  'CONTRACHEQUE', 'LAUDO', 'PLANILHA DE CÁLCULO', 'PLANILHA DE CÁLCULO', 'PLANILHA DE CÁLCULO',
                  'CÁLCULO', 'OUTROS']
        for x in range(1, 19):
            navegador.find_element_by_xpath('//*[@id="txtTipo_' + str(x) + '"]').send_keys(lista2[x - 1])
            sleep(1)
            navegador.find_element_by_xpath('//*[@id="txtTipo_' + str(x) + '"]').send_keys(Keys.ENTER)
            sleep(0.5)
            if x == 6:
                navegador.find_element_by_xpath('//*[@id="txtObservacao_6"]').send_keys('Registro funcional')
            elif x == 7:
                navegador.find_element_by_xpath('//*[@id="txtObservacao_7"]').send_keys('Lei')
            elif x == 8:
                navegador.find_element_by_xpath('//*[@id="txtObservacao_8"]').send_keys('Lei')
            elif x == 9:
                navegador.find_element_by_xpath('//*[@id="txtObservacao_9"]').send_keys('Tabela de reenquadramento')
            elif x == 18:
                navegador.find_element_by_xpath('//*[@id="txtObservacao_18"]').send_keys('Lei')
        if tramit_deficiencia == 'Sim' or tramit_pref == 'Sim':
            navegador.find_element_by_xpath('//*[@id="txtTipo_19"]').send_keys('LAUDO')
            sleep(1)
            navegador.find_element_by_xpath('//*[@id="txtTipo_19"]').send_keys(Keys.ENTER)

    navegador.find_element_by_xpath('//*[@id="btnEnviarArquivos"]').click()  # CONFIRMA SELEÇÃO DE ARQUIVOS
    sleep(5)
    return 1


try:
    WebDriverWait(navegador, 3).until(EC.alert_is_present())
    alert = navegador.switch_to.alert
    alert.accept()
    if erro_upload() == '1':
        while erro_upload() == '1':
            erro_upload()
    classificacao_docs()
except TimeoutException:
    classificacao_docs()

msgbox("Confira os dados e documentos antes de finalizar a distribuição!")

# FINALIZAR DISTRIBUIÇÃO
# CAPTURAR DADOS DO RECIBO
# CADASTRAR PROCESSO NO ESPAIDER
message = "Distribuir inicial de forma automatizada?"
title = "Distribuição automatizada"
if boolbox(message, title, ["Sim", "Não"]):
    finaliza = 'sim'
else:
    finaliza = 'nao'

if finaliza == 'sim':
    sleep(1.5)
    navegador.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div/div/form/div[3]/button[2]').click()
    sleep(1)

    navegador.find_element_by_xpath('//*[@id="sbmConfirmar"]').click()
    sleep(3)

    try:
        navegador.find_element_by_xpath('//*[@id="btnFechar"]').click()
        sleep(2)
    except:
        pyautogui.press('tab', presses=6, interval=0.3)
        pyautogui.press('enter')

    navegador.find_element_by_xpath('//*[@id="btnImprimir"]')
    sleep(1)
    pyautogui.press('esc')
    pyautogui.keyDown('ctrl')
    pyautogui.press('s')
    pyautogui.keyUp('ctrl')
    sleep(1)
    pyautogui.write(documentos + 'Extrato distribuicao')

# botao confirma ajuizamento> //*[@id="sbmConfirmar"]
# fechar guia de custas: 5 tabs ou //*[@id="btnFechar"]
# imprimir extrato: //*[@id="btnImprimir"]
# extrato: 'esc' + ctrl+s

# juizo no espaider: 01ª Vara da Faz. Pub. e Reg. Pub.


# AVISO CONCLUSÃO
ctypes.windll.user32.MessageBoxW(0, "Distribuição finalizada!", "Distribuição Inicial - Blumenau", 1)

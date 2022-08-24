from easygui import msgbox, boolbox
from selenium import webdriver
from time import sleep
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from pathlib import Path
import PyPDF2
import re
import pandas as pd


def clica_xpath(driver, xpath):
    driver.find_element(by=By.XPATH, value=xpath).click()


def erro_upload():
    message = "Erro no upload de alguns arquivos. Resolva-os antes de prosseguir"
    title = "Erro upload"
    if boolbox(message, title, ["Continuar", "Cancelar"]):
        controle_erro = '0'
    else:
        controle_erro = '1'
    return controle_erro


def ler_inicial():
    # Abre o arquivo pdf
    pdf_file = open(documentos + r'\1 Inicial.pdf', 'rb')

    # Faz a leitura usando a biblioteca
    read_pdf = PyPDF2.PdfFileReader(pdf_file, strict=False)

    # lê a primeira página completa
    page = read_pdf.getPage(0)

    # extrai apenas o texto
    page_content = page.extractText()

    # faz a junção das linhas
    parsed = ''.join(page_content)

    # remove as quebras de linha
    parsed = re.sub('n', '', parsed)

    # Tratamento do cpf
    posterior_cpf = parsed.split('CPF')[1]
    virgula_cpf = posterior_cpf.split(',')[0]
    resultado_cpf = virgula_cpf[-14:]
    resultado_cpf = ''.join(resultado_cpf.split('.'))
    resultado_cpf = ''.join(resultado_cpf.split('-'))

    # Tratamento do nome
    posterior_nome = parsed.split(',')[0]
    sc = parsed.split('SC')[0]
    nome = posterior_nome[len(sc) + 6:]
    print('#' * 40)
    print('Nome:', nome)
    print('CPF:', resultado_cpf)

    # lê a penultima página completa
    page = read_pdf.getPage(-2)

    # extrai apenas o texto
    page_content = page.extractText()

    # faz a junção das linhas
    parsed = ''.join(page_content)

    # remove as quebras de linha
    parsed = re.sub('n', '', parsed)

    # Tratamento do valor da causa
    posterior_valor = parsed.split('causa o valor de')[1]
    virgula = posterior_valor.split(',')[0]
    resultado_valor = posterior_valor[4:len(virgula) + 3]
    print('Valor da causa: R$', resultado_valor)

    return resultado_cpf, resultado_valor


# DADOS FIXOS
login = 'RS111059'
senha = 'Alegrianoesforço!'
comarca = 'Blumenau'
cnpj_blumenau = '83108357000115'
oab_mauricio = 'RS036798'
assunto = 'REAJUSTES DE REMUNERAÇÃO, PROVENTOS OU PENSÃO'

# INPUT DE DADOS
servidores = pd.read_excel('Servidores.xlsx')

# NAVEGADOR
navegador = webdriver.Chrome(ChromeDriverManager().install())
navegador.implicitly_wait(60)
actions = ActionChains(navegador)
link = 'https://eproc1g.tjsc.jus.br/eproc/'

# ACESSO AO LOGIN
navegador.get(url=link)

navegador.find_element(by=By.ID, value='txtUsuario').send_keys(login)  # CAMPO LOGIN
navegador.find_element(by=By.ID, value='pwdSenha').send_keys(senha)  # CAMPO SENHA
sleep(0.5)
navegador.find_element(by=By.ID, value='sbmEntrar').click()  # ACESSA PELO LOGIN

# ITERAÇÃO POR LINHA DO ARQUIVO SERVIDORES
for linha in servidores.iterrows():
    cliente = linha[1]['Nome do cliente']
    documentos = rf'O:\{cliente[0]}\{cliente}\Hora-atividade\Cumprimento de sentença'
    cpf, valor_acao = ler_inicial()

    # PRIMEIRA PÁGINA DE PARÂMETROS
    clica_xpath(navegador, '//*[@id="main-menu"]/li[10]/a')  # DISTRIBUIR INICIAL
    Select(navegador.find_element(by=By.ID, value='selLocalidadeJudicial')).select_by_value('8')  # SELECIONA BLUMENAU

    Select(navegador.find_element(by=By.ID, value='selRitoProcesso')).select_by_value('2')  # RITO ORDINÁRIO

    Select(navegador.find_element(by=By.ID, value='selIdGrupoCompetencia')).select_by_value('5')  # FAZENDA PÚBLICA

    Select(navegador.find_element(by=By.ID, value='selIdClasseJudicial')).select_by_value('0000014559')

    navegador.find_element(by=By.ID, value='txtProcessoOriginario').send_keys('03157411320188240008')  # COLETIVA

    navegador.find_element(by=By.ID, value='txtValorCausa').send_keys(valor_acao)

    navegador.find_element(by=By.ID, value='txtUsuario').send_keys(oab_mauricio)  # INCLUSÃO MAURÍCIO
    sleep(10)

    navegador.find_element(by=By.ID, value='txtUsuario').send_keys(Keys.ENTER)
    sleep(1)

    navegador.find_element(by=By.ID, value='btnIncluir').click()
    sleep(1)

    clica_xpath(navegador, '//*[@id="btnSalvar"]')  # PROXIMA PAGINA

    # SEGUNDA PÁGINA DE PARÂMETROS
    navegador.find_element(by=By.XPATH, value='//*[@id="txtFiltroPesquisa"]').send_keys(assunto)  # ASSUNTO
    sleep(0.5)

    clica_xpath(navegador, '//*[@id="btnPesquisar"]')  # BOTAO PESQUISA ASSUNTO
    sleep(4)

    clica_xpath(navegador, '//*[@id="011103_anchor"]/span/span')  # SELECIONA O ASSUNTO
    sleep(0.5)

    clica_xpath(navegador, '//*[@id="btnIncluirAssunto"]')  # INCLUI O ASSUNTO
    sleep(1)

    Select(navegador.find_element(by=By.ID, value='selNumCodCompetencia')).select_by_value('127')

    try:
        navegador.implicitly_wait(4)
        clica_xpath(navegador, '//*[@id="backTop"]')  # PRÓXIMA PÁGINA
        navegador.implicitly_wait(60)
    except:
        pass

    sleep(0.7)
    navegador.find_element(by=By.ID, value='btnSalvar').click()

    # TERCEIRA PÁGINA DE PARÂMETROS
    navegador.find_element(by=By.XPATH, value='//*[@id="txtCpfCnpj"]').send_keys(cpf)  # CAMPO CPF
    sleep(0.5)

    clica_xpath(navegador, '//*[@id="btnConsultarNome"]')  # BOTÃO PARA CONSULTAR CPF
    sleep(1)

    try:
        clica_xpath(navegador, '//*[@id="btnIncluir"]')  # BOTAO INCLUIR AUTOR
        sleep(1)
    except:
        print('Cadastro incompleto')
        sleep(1)

    Select(navegador.find_elements(by=By.TAG_NAME, value='select')[9]).select_by_value('6')
    sleep(0.4)

    navegador.find_element(by=By.XPATH, value='//*[@id="btnProxima"]').click()  # PROXIMA PAGINA
    sleep(1)

    # QUARTA PÁGINA DE PARÂMETROS
    Select(navegador.find_element(by=By.XPATH, value='//*[@id="selTipoPessoa"]')).select_by_value('ENT')
    sleep(0.5)

    Select(navegador.find_element(by=By.XPATH, value='//*[@id="selEntidade"]')).select_by_value(
        '771230778800100040000000000395')  # SELECIONA MUNICÍPIO DE BLUMENAU
    sleep(0.5)

    clica_xpath(navegador, '//*[@id="btnIncluirEnt"]')  # BOTAO DE INCLUIR RÉU
    sleep(0.5)

    clica_xpath(navegador, '//*[@id="btnProxima"]')  # PASSA PARA A PROXIMA PAGINA
    sleep(2)

    # QUINTA PÁGINA DE PARÂMETROS
    # SELECIONA TODOS OS ARQUIVOS PDF DA PASTA
    controle_erro = ''
    cont = 1
    nomes_arquivos = []
    documentos = Path(documentos)
    for i in Path.iterdir(documentos):
        if str(i)[-3:] == 'pdf':
            if cont > 1:
                clica_xpath(navegador, '//*[@id="lblAdicionarDocumento"]')
                sleep(0.4)
            nomes_arquivos.append(i.name)
            navegador.find_element(by=By.XPATH, value=f'//*[@id="_{cont}"]/div/div[2]/input').send_keys(str(documentos)
                                                                                                        + '\\' + i.name)
            cont += 1
    sleep(5)
    print('Arquivos encontrados:', nomes_arquivos)

    # CLASSIFICA OS DOCUMENTOS
    def classificacao_docs():
        # CLASSIFICA TODOS OS ARQUIVOS
        if '8 RG.pdf' in nomes_arquivos:
            lista = ['PETIÇÃO INICIAL', 'CÁLCULO', 'CÁLCULO', 'TÍTULO EXECUTIVO JUDICIAL', 'FICHA FINANCEIRA',
                     'PROCURAÇÃO', 'SUBSTABELECIMENTO', 'IDENTIDADE']
            tipos_com_id = {'PETIÇÃO INICIAL': '1', 'CÁLCULO': '4', 'TÍTULO EXECUTIVO JUDICIAL': '175',
                            'FICHA FINANCEIRA': '39', 'PROCURAÇÃO': '2', 'SUBSTABELECIMENTO': '26', 'IDENTIDADE': '8'}
            for x in range(1, len(lista) + 1):
                js_code = f"document.getElementById('selTipoArquivo_{x}').value = '{tipos_com_id[lista[x - 1]]}'"
                navegador.execute_script(js_code)
                sleep(1)
        else:
            lista = ['PETIÇÃO INICIAL', 'CÁLCULO', 'CÁLCULO', 'TÍTULO EXECUTIVO JUDICIAL', 'FICHA FINANCEIRA',
                     'PROCURAÇÃO', 'SUBSTABELECIMENTO']
            tipos_sem_id = {'PETIÇÃO INICIAL': '1', 'CÁLCULO': '4', 'TÍTULO EXECUTIVO JUDICIAL': '175',
                            'FICHA FINANCEIRA': '39', 'PROCURAÇÃO': '2', 'SUBSTABELECIMENTO': '26'}
            for x in range(1, len(tipos_sem_id) + 1):
                js_code = f"document.getElementById('selTipoArquivo_{x}').value = '{tipos_sem_id[lista[x - 1]]}'"
                navegador.execute_script(js_code)
                sleep(1)

        clica_xpath(navegador, '//*[@id="btnEnviarArquivos"]')  # CONFIRMA SELEÇÃO DE ARQUIVOS
        sleep(5)
        return 1

    # VERIFICA SE EXISTE ERRO DE UPLOAD
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

    msgbox("Confira os dados e documentos antes de finalizar a distribuição. Antes de aceitar este alerta,"
           "finalize a distribuição e aguarde o site voltar para a página inicial.!")

    navegador.maximize_window()

    try:
        clica_xpath(navegador, '//*[@id="backTop"]')
    except:
        pass

    sleep(1.5)

    clica_xpath(navegador, '//*[@id="btnSalvar"]')
    sleep(2)

    iframe = navegador.find_element(by=By.XPATH, value='//*[@id="ifrSubFrm"]')
    navegador.switch_to.frame(iframe)
    clica_xpath(navegador, '//*[@id="sbmConfirmar"]')

    navegador.switch_to.default_content()

    iframe = navegador.find_element(by=By.XPATH, value='//*[@id="ifrSubFrm"]')
    navegador.switch_to.frame(iframe)
    element = navegador.find_element(by=By.XPATH, value='//*[@id="btnFechar"]')
    actions.move_to_element(element).click().perform()
    navegador.switch_to.default_content()

    # clica_xpath(navegador, '//*[@id="btnImprimir"]')

    texto_extrato = navegador.find_element(by=By.XPATH, value='/html/body/script[50]').text
    texto_extrato = texto_extrato.split('processo_extrato_distribuicao&hash=')[1]
    texto_extrato = texto_extrato.split(',')[0]
    print(texto_extrato)

    navegador.get(f'https://eproc1g.tjsc.jus.br/eproc/controlador.php?acao=processo_extrato_distribuicao&hash={texto_extrato}')
    sleep(4)

    with open("page_source.html", "w") as f:
        f.write(navegador.page_source)

    navegador.switch_to.alert.dismiss()
    msgbox('teste')





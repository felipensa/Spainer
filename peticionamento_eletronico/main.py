from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from easygui import *
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from datetime import date
import os
from pathlib import Path
import pyautogui


navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))  # SET NAVEGADOR
navegador.implicitly_wait(5)

# DADOS
loginFelipe = 'RS111059'
senhaFelipe = 'Alegrianoesforço!'

loginMauricio = 'RS036798'
senhaMauricio = 'Barbieri5515*'
data = date.today()
data_em_texto = data.strftime('%d%m%Y')

# INPUT
message = "Qual usuário fará o peticionamento?"
title = "Usuário"
if boolbox(message, title, ["Felipe", "Maurício"]):
    login = loginFelipe
    senha = senhaFelipe
else:
    login = loginMauricio
    senha = senhaMauricio

message = "Qual o grau de jurisdição?"
title = "Jurisdição"
if boolbox(message, title, ["1º Grau", "2º Grau"]):
    url = 'https://eproc1g.tjrs.jus.br/eproc/'
else:
    url = 'https://eproc2g.tjrs.jus.br/eproc/'

message = "Protocolo individual ou em lote?"
title = "Processamento"
if boolbox(message, title, ["Individual", "Lote"]):
    lote = False
else:
    lote = True

if not lote:  # PARA PROTOCOLO INDIVIDUALIZADO
    print('Individual')
    msg = "Informe o número do processo SEM PONTUAÇÃO"
    title = "Dados"
    processo = enterbox(msg, title)


    def acessoEproc():
        navegador.get(url)  # ABRE NAVEGADOR
        navegador.find_element_by_xpath('//*[@id="txtUsuario"]').send_keys(login)  # CAMPO LOGIN
        navegador.find_element_by_xpath('//*[@id="pwdSenha"]').send_keys(senha)  # CAMPO SENHA
        navegador.find_element_by_xpath('//*[@id="sbmEntrar"]').click()  # BOTAO ENTRAR
        msgbox("Resolva o Captcha caso apareça!")
        navegador.find_element_by_xpath('//*[@id="tr0"]').click()  # SELECIONA PERFIL ADVOGADO
        sleep(1.5)
        return 1

    acessoEproc()

    processo = str(processo)
    print('-' * 50)
    print(processo)

    navegador.find_element_by_xpath('//*[@id="navbar"]/div/div[3]/div[4]/form/input[1]').send_keys(
        processo)  # ENVIA NUMERO PROCESSO PARA CAMPO PESQUISA
    navegador.find_element_by_xpath(
        '//*[@id="navbar"]/div/div[3]/div[4]/form/button[1]').click()  # BOTAO PARA PESQUISAR
    botoes = navegador.find_elements_by_css_selector('a.infraButton')

    # SELECIONAR EVENTO
    msg = "Qual evento será lançado no EPROC?"
    title = "Evento EPROC"
    choices = [
        'ACORDO DE NAO-PERSECUÇAO PENAL',
        'AGRAVOS DE DECISÃO DENEGATÓRIA DE REC. ESPECIAL E EXT.',
        'ALEGAÇÕES FINAIS',
        'APELAÇÃO',
        'APRESENTAÇÃO DE QUESITOS',
        'CIÊNCIA, COM RENÚNCIA AO PRAZO',
        'COMUNICAÇÕES',
        'CONTESTAÇÃO',
        'CONTRARRAZOES',
        'DEFESA PRÉVIA',
        'DENÚNCIA',
        'EMBARGOS À AÇÃO MONITÓRIA',
        'EMBARGOS DE DECLARAÇÃO',
        'EMBARGOS INFRINGENTES',
        'EXCEÇÃO DE PRÉ-EXECUTIVIDADE',
        'EXECUÇÃO/CUMPRIMENTO DE SENTENÇA',
        'GUIAS DE RECOLHIMENTO / DEPÓSITO / CUSTAS',
        'IMPUGNAÇÃO AO CUMPRIMENTO DE SENTENÇA',
        'IMPUGNAÇÃO AOS EMBARGOS',
        'INCIDENTE DE UNIFORMIZAÇÃO DE JURISPRUDÊNCIA',
        'LAUDO COMPLEMENTAR',
        'LAUDO PERICIAL',
        'MANIFESTAÇÃO (ART. 402 CPP)',
        'MEMORIAIS',
        'OFÍCIO',
        'PARECER',
        'PETIÇÃO',
        'PETIÇÃO - ADITAMENTO À DENÚNCIA',
        'PETIÇÃO - EMENDA A INICIAL',
        'PETIÇÃO - IMPUGNAÇÃO AOS CÁLCULOS',
        'PETIÇÃO - PEDIDO DE INSERÇÃO DE RESTRIÇÃO NO SERASAJUD',
        'PETIÇÃO - PEDIDO DE LIMINAR/ANTECIPAÇÃO DE TUTELA',
        'PETIÇÃO - PEDIDO DE RECONSIDERAÇÃO',
        'PETIÇÃO - PEDIDO DE RETIRADA DE RESTRIÇÃO NO SERASAJUD',
        'PETIÇÃO - PEDIDO DE SUSPENSÃO CONDICIONAL DO PROCESSO',
        'PETIÇÃO - PEDIDO DE TRANSAÇÃO PENAL',
        'PETIÇÃO - PRIORIDADE DE TRAMITAÇÃO',
        'PROCURAÇÃO',
        'RAZÕES DE APELAÇÃO CRIMINAL',
        'RAZÕES DE RECURSO EM SENTIDO ESTRITO',
        'RECONVENÇÃO',
        'RECURSO ADESIVO',
        'RECURSO ESPECIAL',
        'RECURSO EXTRAORIDNÁRIO',
        'RENÚNCIA AO PRAZO',
        'RÉPLICA',
        'RESPOSTA'
    ]
    choice = choicebox(msg, title, choices)

    for i in botoes:
        if i.text == "Movimentar/Peticionar":
            i.click()
            break
        else:
            pass

    navegador.find_element_by_xpath('//*[@id="lblListarEvento"]').click()  # CLICA EM LISTAR TODOS
    navegador.find_element_by_xpath('//*[@id="txtEvento"]').send_keys(choice)  # ESCREVE O EVENTO NO CAMPO
    sleep(1)
    navegador.find_element_by_xpath('//*[@id="txtEvento"]').send_keys(Keys.ENTER)  # DÁ UM ENTER PRA SALVAR O EVENTO
    sleep(2)
    navegador.find_element(by=By.XPATH, value=f'//*[@id="{processo}_1"]/div/div[2]').click()
    sleep(2)

    def ordenacao():
        p = Path(f'./PROTOCOLO/{processo} - {data_em_texto}')
        arquivos = [x.name for x in p.iterdir() if x.is_file()]
        arquivos_format = []
        for x in arquivos:
            arquivos_format.append('"' + x + '"')
        id = []
        dados = []
        for item in arquivos:
            parte = item.split('-')
            id.append(parte[0])
        for numero in id:
            indice = id.index(numero)
            id[indice] = int(numero)
            dados.append([arquivos_format[indice], id[indice]])
        dados_ordenados = []
        indice_dados = sorted(id)
        for elemento in indice_dados:
            for item in dados:
                if elemento == item[1]:
                    dados_ordenados.append(item[0])
        print(f'Resultado: {dados_ordenados}')
        dados_ordenados_str = ' '.join(dados_ordenados)
        cwd = Path.cwd()
        full_path = cwd.joinpath(p)
        full_path = str(full_path)
        return full_path, dados_ordenados_str
    ordenacao()


    def anexa_documentos():
        pyautogui.write(ordenacao()[0])
        pyautogui.press('enter')
        sleep(1.5)
        pyautogui.write(ordenacao()[1])
        sleep(1)
        pyautogui.press('enter')
    anexa_documentos()

msgbox('Processo finalizado. Por favor, classifique os documentos e protocole a peça')

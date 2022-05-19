from time import sleep
import xlrd
from easygui import*
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

# DADOS
login = 'RS036798'
senha = 'Barbieri5515*'
url = 'https://eproc2g.tjrs.jus.br/eproc/'

# WEBDRIVER
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
    'download.prompt_for_download': False,
    'plugins.always_open_pdf_externally': True,
})
navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)
navegador.implicitly_wait(15)

# COLETA DE DADOS
wb = xlrd.open_workbook('intimacoes.xls', encoding_override='iso-8859-1')
ws = wb.sheet_by_name('Processos Pendentes - Urgente')
linhas = ws.nrows
colunas = ws.ncols

# ACESSO AO EPROC
navegador.get(url)  # ABRE NAVEGADOR
navegador.find_element(by=By.XPATH, value='//*[@id="txtUsuario"]').send_keys(login)  # CAMPO LOGIN
navegador.find_element(by=By.XPATH, value='//*[@id="pwdSenha"]').send_keys(senha)  # CAMPO SENHA
navegador.find_element(by=By.XPATH, value='//*[@id="sbmEntrar"]').click()  # BOTO ENTRAR
msgbox("Resolva o Captcha caso apareça!")
navegador.find_element(by=By.XPATH, value='//*[@id="tr0"]').click()  # SELECIONA PERFIL ADVOGADO
sleep(1.5)

# LISTAS
cj_id_intimacoes = []
cj_processos = []
cj_originarios = []
cj_precatorios = []
cj_erros = []
cj_datas = []
cj_requerentes = []
cj_requeridos = []
cj_clientes = []
cj_destinatarios = []
cj_descricoes = []
cj_intimacoes = []
cj_eventos = []
cj_notas = []
id_intimacao = 100

# ITERAÇÃO
for item in range(2, linhas):
    processo = str(ws.cell_value(rowx=item, colx=0))
    print('x' * 50)
    print(processo)
    navegador.find_element(by=By.XPATH,
                           value='//*[@id="navbar"]/div/div[3]/div[4]/form/input[1]').send_keys(processo)   # ENVIA NUMERO DO PROCESSO PARA CAMPO PESQUISA
    navegador.find_element(by=By.XPATH,
                           value='//*[@id="navbar"]/div/div[3]/div[4]/form/button[1]').click()  # BOTAO PARA PESQUISAR

    # DADOS DE CAPA
    # TENTA COLETAR PROCESSO ORIGINÁRIO
    try:
        processo_originario_1 = navegador.find_element(by=By.XPATH,
                                                       value='//*[@id="tableRelacionado"]/tbody/tr[1]/td[1]').text
        if len(processo_originario_1) >= 20:
            processo_originario_1 = processo_originario_1[0:-3]
        print('Pegou primeiro originário...')
        processo_originario_2 = navegador.find_element(by=By.XPATH,
                                                       value='//*[@id="tableRelacionado"]/tbody/tr[2]/td[1]').text
        if len(processo_originario_2) >= 20:
            processo_originario_2 = processo_originario_2[0:-3]
        print('Pegou segundo originário...')
        processo_originario = 'OK'
    except NoSuchElementException:
        processo_originario = 'Sem dados de processo originário'
        print('Sem dados de originário')

    sleep(1)
    data_distribuicao = navegador.find_element(by=By.XPATH,
                                               value='//*[@id="txtAutuacao"]').text  # COLETA DATA DE DISTRIBUIÇÃO

    sleep(0.5)
    data_reduzida = data_distribuicao[0:10]

    navegador.implicitly_wait(3)
    try:
        navegador.find_element(by=By.XPATH, value='//*[@id="carregarOutrosA"]').click()   # ABRE TODAS AS PARTES
        sleep(2)
    except NoSuchElementException:
        pass
    navegador.implicitly_wait(15)

    nomes_das_partes = navegador.find_elements(by=By.CLASS_NAME,
                                               value='infraNomeParte')  # COLETA NOME DE TODAS AS PARTES
    lista_requerentes = []

    for parte in nomes_das_partes:
        lista_requerentes.append(parte.text)
    requerente = lista_requerentes[0]
    requerido = navegador.find_element(by=By.XPATH, value='//*[@id="spnNomeParteReu0"]').text

    lista_td_partes = []
    trs = len(navegador.find_elements(by=By.XPATH, value='//*[@id="tblPartesERepresentantes"]/tbody/tr'))
    for tr in range(2, trs+1):
        for td in range(1, 3):
            print(f'TR: {tr} - TD: {td}')
            try:
                lista_td_partes.append(navegador.find_element(by=By.XPATH,
                                                              value=f'//*[@id="tblPartesERepresentantes"]/tbody/tr[{tr}]/td[{td}]').text)
            except NoSuchElementException:
                pass
    print(f'QUANTIDADE DE TRS: {trs}')
    print(f'PARTES: {lista_td_partes}')
    print(lista_requerentes)
    cliente = []
    for td in lista_td_partes:
        for requerente in lista_requerentes:
            if requerente in td and 'MAURÍCIO LINDENMEYER BARBIERI' in td:
                cliente.append(requerente)
    print(f'CLIENTES: {cliente}')

    # COLETA INTIMAÇÕES
    controle_intimacoes_clientes = 0

    controle_clara = navegador.find_elements(by=By.CSS_SELECTOR,
                                             value='[class="infraTrClara infraEventoPrazoAguardando"]')  # CHECA QUANTOS ELEMENTOS VERMELHOS TEM NA PARTE CLARA
    controle_escura = navegador.find_elements(by=By.CSS_SELECTOR,
                                              value='[class="infraTrEscura infraEventoPrazoAguardando"]')  # CHECA QUANTOS ELEMENTOS VERMELHOS TEM NA PARTE ESCURA
    controle_clara_amarelo = navegador.find_elements(by=By.CSS_SELECTOR,
                                                     value='[class="infraTrClara infraEventoPrazoAberto"]')  # CHECA QUANTOS ELEMENTOS AMARELOS TEM NA PARTE CLARA
    controle_escura_amarelo = navegador.find_elements(by=By.CSS_SELECTOR,
                                                      value='[class="infraTrEscura infraEventoPrazoAberto"]')  # CHECA QUANTOS ELEMENTOS AMARELOS TEM NA PARTE ESCURA

    print(f'XXX CONTROLES XXX'
          f'\nCLARA-VERMELHA: {len(controle_clara)};'
          f'\nESCURA-VERMELHA: {len(controle_escura)};'
          f'\nCLARA-AMARELA: {len(controle_clara_amarelo)};'
          f'\nESCURA-AMARELA: {len(controle_escura_amarelo)}')

    destinarios = ''
    conjunto_vermelhos = []
    conjunto_amarelos = []

    for elemento_claro in controle_clara:
        conjunto_vermelhos.append(elemento_claro.text)
    for elemento_escuro in controle_escura:
        conjunto_vermelhos.append(elemento_escuro.text)
    for elemento_claro_amarelo in controle_clara_amarelo:
        conjunto_amarelos.append(elemento_claro_amarelo.text)
    for elemento_escuro_amarelo in controle_escura_amarelo:
        conjunto_amarelos.append(elemento_escuro_amarelo.text)
    print(f'VERMELHOS: {conjunto_vermelhos}')
    print(f'AMARELOS: {conjunto_amarelos}')

    # CONTROLE VERMELHO - PRAZOS PENDENTES
    if len(controle_clara) > 0 or len(controle_escura) > 0:
        print("Prazo Aguardando Abertura")
        for vermelho in conjunto_vermelhos:
            for pessoa in cliente:
                if pessoa in vermelho:
                    print(f'PESSOA INTIMADA: {pessoa}')
                    destinarios += pessoa + ';'
                    controle_intimacoes_clientes += 1
                    if controle_intimacoes_clientes < 1:
                        indice_intimacao = conjunto_vermelhos.index(vermelho)
                        print(f'INDICE VERMELHO: {indice_intimacao}')
                    else:
                        print('ERRO NO INDICE VERMELHO...')
                    controle_intimacoes_clientes += 1

        lista_destinatarios = destinarios.split(';')
        print(f'LISTA DESTINATÁRIOS: {lista_destinatarios}')
        total_intimacoes = controle_intimacoes_clientes / len(conjunto_vermelhos)
        print(f'TOTAL INTIMAÇÕES: {total_intimacoes}')

        conjunto_vermelhos[indice_intimacao] = conjunto_vermelhos[indice_intimacao].replace('\n', '')
        texto_posterior = conjunto_vermelhos[indice_intimacao].split(':')
        print(f'POSTERIOR: {texto_posterior}')
        texto_anterior = conjunto_vermelhos[indice_intimacao].split('(')  # menos um do posterior
        print(f'ANTERIROR: {texto_anterior}')
        indice_parenteses = texto_posterior[2].find('(')
        print(f'PARENTESES: {indice_parenteses}')
        resultado_primeiro = texto_posterior[2][0:indice_parenteses]  # UNIÃO DOS RESULTADOS
        print(f'Resultado 1: {resultado_primeiro}')
        # resultado_segundo = resultado_primeiro.split(' ')
        # print(f'Resultado 2: {resultado_segundo}')

        resultado_segundo = resultado_primeiro[-2:]
        try:
            resultado = int(resultado_segundo)
            resultado = str(resultado)
            print('Resultado com dois algarismos')
        except:
            resultado = resultado_primeiro[-1:]
            print('Resultado com um algarismo')

        # if len(resultado_primeiro) > 2:
        #     resultado = resultado_segundo[1]
        #     print(f'RESULTADO: {resultado}')
        # else:
        #     resultado = resultado_primeiro.replace(' ', '')
        #     print(f'RESULTADO ELSE: {resultado}')

        print('RESULTADO: \n', resultado)
        navegador.find_element(by=By.CSS_SELECTOR, value=f'[id="trEvento{resultado}"]')
        descricao = navegador.find_element(by=By.XPATH, value=f'//*[@id="trEvento{resultado}"]/td[3]/label').text
        cj_descricoes.append(descricao)
        # um evento só
        descricao_nota = navegador.find_element(by=By.XPATH, value=f'//*[@id="trEvento{resultado}"]/td[3]/label').text
        print('Descrição da Nota: \n', descricao_nota)


    # CONTROLE AMARELO - PRAZOS ABERTOS
    elif len(controle_clara_amarelo) > 0 or len(controle_escura_amarelo) > 0:
        print("Prazo ABERTO")
        for amarelo in conjunto_amarelos:
            for pessoa in cliente:
                if pessoa in amarelo:
                    print(f'PESSOA INTIMADA: {pessoa}')
                    destinarios += pessoa + ';'
                    if controle_intimacoes_clientes < 1:
                        indice_intimacao = conjunto_amarelos.index(amarelo)
                        print(f'INDICE AMARELO: {indice_intimacao}')
                    else:
                        print('ERRO NO INDICE AMARELO...')
                    controle_intimacoes_clientes += 1

        lista_destinatarios = destinarios.split(';')
        print(f'LISTA DESTINATÁRIOS: {lista_destinatarios}')
        total_intimacoes = controle_intimacoes_clientes / len(conjunto_amarelos)
        print(f'TOTAL INTIMAÇÕES: {total_intimacoes}')

        conjunto_amarelos[indice_intimacao] = conjunto_amarelos[indice_intimacao].replace('\n', '')
        texto_posterior = conjunto_amarelos[indice_intimacao].split(':')
        print(f'POSTERIOR: {texto_posterior}')
        texto_anterior = conjunto_amarelos[indice_intimacao].split('(')  # menos um do posterior
        print(f'ANTERIROR: {texto_anterior}')
        indice_parenteses = texto_posterior[2].find('(')
        print(f'PARENTESES: {indice_parenteses}')
        resultado_primeiro = texto_posterior[2][0:indice_parenteses]  # UNIÃO DOS RESULTADOS
        print(f'Resultado 1: {resultado_primeiro}')
        # resultado_segundo = resultado_primeiro.split(' ')
        # print(f'Resultado 2: {resultado_segundo}')

        resultado_segundo = resultado_primeiro[-2:]
        try:
            resultado = int(resultado_segundo)
            resultado = str(resultado)
            print('Resultado com dois algarismos')
        except:
            resultado = resultado_primeiro[-1:]
            print('Resultado com um algarismo')

        # if len(resultado_primeiro) > 2:
        #     resultado = resultado_segundo[1]
        #     print(f'RESULTADO: {resultado}')
        # else:
        #     resultado = resultado_primeiro.replace(' ', '')
        #     print(f'RESULTADO ELSE: {resultado}')

        print('RESULTADO: \n', resultado)
        navegador.find_element(by=By.CSS_SELECTOR, value=f'[id="trEvento{resultado}"]')
        descricao = navegador.find_element(by=By.XPATH, value=f'//*[@id="trEvento{resultado}"]/td[3]/label').text
        cj_descricoes.append(descricao)
        print('Descrição da Nota: \n', descricao)

        #COLETA EVENTO
        try:
            navegador.find_element(by=By.XPATH,
                               value=f'//*[@id="trEvento{resultado}"]/td[5]/a').click()  # ABRE O ÚLTIMO DOCUMENTO(EVENTO)
            sleep(1)
            navegador.switch_to.window(navegador.window_handles[1])  # TROCA DE GUIA

            paragrafo = navegador.find_elements(by=By, value='paragrafoPadrao')
            contLinha = len(paragrafo)
            x = 0
            paragrafoLinhas = []
            for i in range(0, contLinha):
                linha = paragrafo[x].text
                x += 1
                paragrafoLinhas.append(linha)
            paragrafoUnificado = "".join(paragrafoLinhas)
        except:
            paragrafoUnificado = 'Sem documentos'

        sleep(1)
        navegador.close()
        navegador.switch_to.window(navegador.window_handles[0])

    # ARMAZENAMENTO DE VALORES
    cj_id_intimacoes.append(id_intimacao)
    cj_processos.append(processo)
    cj_originarios.append(processo_originario_1)
    cj_precatorios.append(processo_originario_2)
    cj_erros.append(processo_originario)
    cj_datas.append(data_reduzida)
    cj_requerentes.append(lista_requerentes[0])
    cj_requeridos.append(requerido)
    cj_clientes.append(cliente)
    cj_eventos.append(resultado)
    cj_notas.append(paragrafoUnificado)
    cj_destinatarios.append(lista_destinatarios)

    id_intimacao += 10
    print(f'Contagem: {item-1}/{linhas}')


dicionario = {'ID': cj_id_intimacoes,
              'Processo': cj_processos,
              'Originário/Precatório': cj_originarios,
              'Precatório/Originário': cj_precatorios,
              'Sem dados do originário': cj_erros,
              'Data distribuição': cj_datas,
              'Primeiro requerente': cj_requerentes,
              'Requerido': cj_requeridos,
              'Clientes': cj_clientes,
              'Descrição': cj_descricoes,
              'Evento ref': cj_eventos,
               'Intimação': cj_notas,
              # 'Destinatários':
              }      #ARMAZENA OS VALORES EM UMA MATRIZ
df = pd.DataFrame(dicionario)        #CRIA O DATAFRAME
df.to_excel('./Lista de notas.xlsx')        #CRIA O ARQUIVO EXCEL





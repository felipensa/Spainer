from time import sleep
import xlrd
from selenium import webdriver
import pandas as pd
from easygui import*
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

# DADOS
login_espaider = 'teste'
senha_espaider = 'Sumerios0394&'
loginFelipe = 'RS111059'
senhaFelipe = 'Alegrianoesforço!'
loginMauricio = 'RS036798'
senhaMauricio = 'Barbieri5515*'
espaider = 'http://barbieriadvogados.dyndns.org:40400/Barbieri/'

# COLETA DADOS EXCEL
wb = xlrd.open_workbook('INTIMACOES.xls')
planilha = wb.sheet_by_name('Planilha1')
total_linhas = planilha.nrows
total_colunas = planilha.ncols

# PARÂMETROS
message_grau = "Qual grau de jurísdição?"
title_grau = "Jurisdição"
if boolbox(message_grau, title_grau, ["1º Grau", "2º Grau"]):
  url = 'https://eproc1g.tjrs.jus.br/eproc/'
else:
  url = 'https://eproc2g.tjrs.jus.br/eproc/'

message_usu = "Qual usuário?"
title_usu = "Usuário"
if boolbox(message_usu, title_usu, ["Felipe", "Maurício"]):
  login = loginFelipe
  senha = senhaFelipe
  socio = 'nao'
else:
  login = loginMauricio
  senha = senhaMauricio
  socio = 'sim'

# ACESSO EPROC
navegador = webdriver.Chrome(ChromeDriverManager().install())
navegador.implicitly_wait(5)
navegador.get(url)  # ABRE NAVEGADOR
navegador.find_element(by=By.XPATH, value='//*[@id="txtUsuario"]').send_keys(login)       # CAMPO LOGIN
navegador.find_element(by=By.XPATH, value='//*[@id="pwdSenha"]').send_keys(senha)     # CAMPO SENHA
navegador.find_element(by=By.XPATH, value='//*[@id="sbmEntrar"]').click()     # BOTAO ENTRAR
msgbox("Resolva o Captcha caso apareça!") # ESPERA O CAPTCHA
sleep(1)
navegador.find_element(by=By.XPATH, value='//*[@id="tr0"]').click()       # SELECIONA PERFIL ADVOGADO

# TABELAS
dados = []
precatorios = []
processos = []
data_dist = []
comarcas = []
requerentes = []
requeridos = []
notas = []
textoVermelho = []
cj_descricao = []

# PESQUISA PROCESSOS PARA CADA LINHA DO EXCEL
for i in range(2, total_linhas):
  processo = str(planilha.cell_value(rowx=i, colx=0))   # DECLARA O PROCESSO
  print('x'*70)
  print(processo)
  navegador.find_element(by=By.XPATH, value='//*[@id="navbar"]/div/div[3]/div[4]/form/input[1]').send_keys(processo)  # ENVIA NUMERO PROCESSO PARA CAMPO PESQUISA
  navegador.find_element(by=By.XPATH, value='//*[@id="navbar"]/div/div[3]/div[4]/form/button[1]').click() # BOTAO PARA PESQUISAR
  sleep(1)

  try:
    processo_originario = navegador.find_element(by=By.XPATH, value='//*[@id="tableRelacionado"]/tbody/tr/td[1]/font/a').text     # COLETA NÚMERO DO PRIMEIRO PROCESSO ORIGINÁRIO
    sleep(0.5)
  except NoSuchElementException:
    processo_originario = 'Sem dados de processo originário'

  # DADOS DA CAPA DO PROCESSO
  data = navegador.find_element(by=By.XPATH, value='//*[@id="txtAutuacao"]').text   # COLETA DATA DE DISTRIBUIÇÃO
  data_reduc = data[0:10]
  try:
    parte_autora = navegador.find_element(by=By.XPATH, value="//a[@data-parte='AUTOR']").text
  except:
    parte_autora = navegador.find_element(by=By.XPATH, value="//span[@data-parte='AUTOR']").text # COLETA OS NOMES DOS REQUERENTES

  try:
    parte_re = navegador.find_element(by=By.XPATH, value="//a[@data-parte='REU']").text
  except:
    parte_re = navegador.find_element(by=By.XPATH, value="//span[@data-parte='REU']").text # COLETA OS NOMES DOS REQUERIDOS

  print(parte_autora)
  print(parte_re)
  if parte_autora == 'MAURICIO DAL AGNOL':
    cliente = parte_re
  else:
    cliente = parte_autora
  sleep(0.5)

  # ARMAZENA VALORES
  dados.append(processo_originario)  # JOGA O VALOR EXTRAIDO PARA A LISTA
  processos.append(processo)  # JOGA O VALOR DO PROCESSO PARA A LISTA
  data_dist.append(data_reduc)  # JOGA VALOR DATA PARA A LISTA
  requerentes.append(parte_autora)
  requeridos.append(parte_re)

  # CONTROLE DE LINHA
  sleep(1)
  controleClara = len(navegador.find_elements(by=By.CSS_SELECTOR, value='[class="infraTrClara infraEventoPrazoAguardando"]'))   # CHECA QUANTOS ELEMENTOS VERMELHOS TEM NA PARTE CLARA
  controleEscura = len(navegador.find_elements(by=By.CSS_SELECTOR, value='[class="infraTrEscura infraEventoPrazoAguardando"]')) # CHECA QUANTOS ELEMENTOS VERMELHOS TEM NA PARTE ESCURA

  controleClaraAmarelo = len(navegador.find_elements(by=By.CSS_SELECTOR, value='[class="infraTrClara infraEventoPrazoAberto"]'))   # CHECA QUANTOS ELEMENTOS AMARELOS TEM NA PARTE CLARA
  controleEscuraAmarelo = len(navegador.find_elements(by=By.CSS_SELECTOR, value='[class="infraTrEscura infraEventoPrazoAberto"]'))  # CHECA QUANTOS ELEMENTOS AMARELOS TEM NA PARTE ESCURA
  print(f'Controle Escura: {controleEscura}; Controle Clara: {controleClara}')

  if controleClaraAmarelo > 0 or controleEscuraAmarelo > 0:
    print('Prazo em aberto')
    processo_originario = 'Prazo em aberto'     # JOGA O VALOR EXTRAIDO PARA A LISTA
    processo = 'Prazo em aberto'     # JOGA O VALOR DO PROCESSO PARA A LISTA
    data_reduc = 'Prazo em aberto'    # JOGA VALOR DATA PARA A LISTA
    precatorio_cert = 'Prazo em aberto'
    tipo_acao = 'Prazo em aberto'
    paragrafoUnificado = 'Prazo em aberto'
  elif controleClara > 0 or controleEscura > 0:
    vermelhoClara = navegador.find_elements(by=By.CSS_SELECTOR, value='[class="infraTrClara infraEventoPrazoAguardando"]')
    vermelhoEscura = navegador.find_elements(by=By.CSS_SELECTOR, value='[class="infraTrEscura infraEventoPrazoAguardando"]')
    vermelho = vermelhoClara + vermelhoEscura
    print(f'VERMELHO: {len(vermelho)}')
    for x in range(0, len(vermelho)):
      textoVermelho.append(vermelho[x].text)
      if cliente in vermelho[x].text:
        controleTexto = x
        print(f'CONTROLE TEXTO: {controleTexto}')
        break

    print(f'TEXTO VERMELHO: {textoVermelho}')
    textoVermelho[controleTexto] = textoVermelho[controleTexto].replace('\n', '')        # TIRA AS QUEBRAS DE LINHA
    posterior = textoVermelho[controleTexto].split(':')                                   # ARMAZENAMENTO DO PRIMEIRO TERMO ATÉ O PRÓXIMO
    anterior = textoVermelho[controleTexto].split('(')                                    # menos um do posterior
    indiceParenteses = posterior[2].find('(')
    resultado1 = posterior[2][0:indiceParenteses]                                               # DOS RESULTADOS
    resultado2 = resultado1.split(' ')

    resultado = resultado2[-1]
    # Por enquanto o código pega o último evento referenciado na intimação. Posteriormente devemos corrigir isso para pegar os dois eventos.
    print('RESULTADO:', resultado)
    sleep(1)
    descricao = navegador.find_element(by=By.XPATH, value=f'//*[@id="trEvento{resultado}"]/td[3]/label').text
    cj_descricao.append(descricao)
    print('DESCRIÇÃO DA NOTA:', descricao)

    if descricao == "Juntada de íntegra do processo":
      paragrafoUnificado = 'Digitalização de Processo'
    else:
      paragrafoUnificado = 'Outro processo'
  else:
    print('Prazo fechado ou erro')
  notas.append(paragrafoUnificado)

  textoVermelho = []
  linha_2 = i -1
  print(f'Contagem: {linha_2}/{total_linhas-2}')


df = {'Principal': processos, 'Originário': dados, 'Data distribuição': data_dist, 'Requerente': requerentes, 'Requerido': requeridos, 'Certidão': notas}      # ARMAZENA OS VALORES EM UMA MATRIZ
df1 = pd.DataFrame(df)        # CRIA O DATAFRAME
df1.to_excel('./Lista de notas.xlsx')        # CRIA O ARQUIVO EXCEL

print('FINALIZADO')
navegador.quit()

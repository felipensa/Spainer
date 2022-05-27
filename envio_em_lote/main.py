import win32com.client as win32
import pandas as pd
from time import sleep

relatorio = pd.read_excel('teste.xlsx')

# CRIA INTEGRAÇÃO COM OUTLOOK
outlook = win32.Dispatch('outlook.application')


for indice in range(len(relatorio['Cliente'])):
    # CRIA UM EMAIL
    email = outlook.CreateItem(0)

    cliente = relatorio.loc[indice, 'Cliente']
    email_cliente = relatorio.loc[indice, 'Email']
    print(cliente, email_cliente)

    # CONFIGURA INFORMAÇÕES DO EMAIL
    email.To = email_cliente
    email.Subject = "Aviso - Barbieri Advogados"
    email.HTMLBody = f"""
    <p>Em resposta ao contato de alguns clientes, informamos que a Barbieri Advogados está ciente da ocorrência de fatos 
    maliciosos utilizando o nome do nosso escritório.</p>
    <p></p>
    <p>A prática consiste no envio de mensagem por Whatsapp por parte de pessoas não identificadas e que não pertencem ao 
    escritório, na tentativa de obter valores de maneira ilícita.</p>
    <p></p>
    <p>À oportunidade, relembramos que a Barbieri Advogados não solicita depósito de valores e não envia boletos 
    bancários com esse tipo de finalidade</p>
    <p></p>
    <p>Frente ao ocorrido, informamos, por fim, que já foram tomadas as providências cabíveis junto às autoridades 
    competentes</p>
    <p></p>
    <p></p>
    <p>Att.</p>
    <p>Barbieri Advogados</p>
    <p></p>
    <br>Praça da Alfândega, 12 - 12º e 13º andares.</br>
    <br>Edifício London Bank - Centro Histórico</br>
    <br>Porto Alegre - Rio Grande do Sul - Brasil.</br>
    <br>CEP 90010-150</br>
    <p></p>
    <br>Telefones: +55 (51) 3224.0169 - 3224.9163</br>
    <br>www.barbieriadvogados.com</br>
    <img src="C:/Users/Administrador/Documents/GitHub/Spainer/envio_em_lote/logo.jpg"></img>
    """

    email.Send()
    sleep(5)
    print(f'Email de {cliente} Enviado')

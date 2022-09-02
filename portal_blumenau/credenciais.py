import pandas as pd
from pathlib import Path
from easygui import *


def consulta_credenciais(caminho_navarro, pessoa):
    print('Consultando credenciais do portal da transparência...')
    pasta_parametros = Path(caminho_navarro + r'\PARAMETROS')
    credenciais = pd.read_excel(str(pasta_parametros) + r'\relatorio_vinculos.xlsx')
    matricula = credenciais[credenciais.Cliente == pessoa].Matricula.values[0]
    login = credenciais[credenciais.Cliente == pessoa].Login.values[0]
    senha = credenciais[credenciais.Cliente == pessoa].Senha.values[0]

    while 1:
        if matricula != '' and login != '' and senha != '':
            break
        else:
            msg = "Informe as credenciais do/a cliente"
            title = "Credenciais"
            nome_campos = ["Matrícula: ", "Login: ", "Senha: "]
            valor_campos = []  # CRIA VETOR PARA RECEBER RESPOSTAS
            valor_campos = multenterbox(msg, title, nome_campos)

            # GARANTE QUE NENHUM CAMPO FIQUE VAZIO
            while 1:
                if valor_campos is None:
                    break
                errmsg = ""
                for i in range(len(nome_campos)):
                    if valor_campos[i].strip() == "":
                        errmsg = errmsg + ('"%s" é um campo obrigatório.\n\n' % nome_campos[i])
                if errmsg == "":
                    break  # no problems found
                valor_campos = multenterbox(errmsg, title, nome_campos, valor_campos)
    print('Credenciais obtidas.')

    return matricula, login, senha

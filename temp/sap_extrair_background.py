# Importing the Libraries
import win32com.client
import sys
import subprocess
import time
import json
from datetime import datetime, timedelta
import traceback
import os

# This function will Login to SAP from the SAP Logon window

SapGuiAuto = ""
application = ""
connection = ""
session = ""


def sap_login():
    global SapGuiAuto
    global application
    global connection
    global session

    try:
        path = r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(10)

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
        connection = application.OpenConnection(
            r"02 PEP - SAP S/4HANA Produção (SAP SCRIPT)", True)
        # r"S4H – Ambiente de Homologação do S4HANA", True)

        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        session.findById("wnd[0]").sendVKey(0)

    except:
        print(sys.exc_info()[0])

    # finally:
    #    session = None
    #    connection = None
    #    application = None
    #    SapGuiAuto = None


def sap_logoff():
    global SapGuiAuto
    global application
    global connection
    global session

    try:
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    except:
        pass
    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


# Entra na transação SM37 e salva os arquivos txt
def salvar_arquivo_background(dados_json):
    global SapGuiAuto
    global application
    global connection
    global session
    #SapGuiAuto = win32com.client.GetObject('SAPGUI')
    #application = SapGuiAuto.GetScriptingEngine
    #connection = application.Children(0)
    #session = connection.Children(0)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nSM37"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    pos_primeira_linha_tabela = 0

    # PROCURAR A PRIMEIRA LINHA DA TABELA(CABEÇALHO)
    for linha in range(1, 100):  # PERCORRER DA LINHA 1 ATE 100 DA TELA
        for coluna in range(1, 20):  # POSICÃO
            if session.findById(f"wnd[0]/usr/lbl[{coluna},{linha}]", False) is not None:
                if session.findById(f"wnd[0]/usr/lbl[{coluna},{linha}]").text == "NomeJob":
                    pos_primeira_linha_tabela = linha
                    break
        else:
            continue
        break

    # PROCURAR A POSICAO INICIAL DAS 3 PRIMEIRAS COLUNAS NA TABELA  - | NomeJob | Status | Hora iníc.planej |
    for coluna in range(1, 200):
        if session.findById(f"wnd[0]/usr/lbl[{coluna},{pos_primeira_linha_tabela}]", False) is not None:
            if session.findById(f"wnd[0]/usr/lbl[{coluna},{pos_primeira_linha_tabela}]").text == "NomeJob":
                pos_nome_job = coluna
            elif session.findById(f"wnd[0]/usr/lbl[{coluna},{pos_primeira_linha_tabela}]").text == "Status":
                pos_status = coluna
            elif session.findById(f"wnd[0]/usr/lbl[{coluna},{pos_primeira_linha_tabela}]").text == "Hora iníc.planej.":
                pos_hora_inicio = coluna

    # PERCORRER AS LINHAS DA TABELA NA POSICAO DAS COLUNAS - | NomeJob | Status | Hora iníc.planej |
    for linha in range(1, 200):
        nomeJob = ""
        horaProg = ""
        jobConcluido = False
        # VERIFICAR SE É UMA LINHA DA TABELA DE JOBS
        if session.findById(f"wnd[0]/usr/lbl[{pos_nome_job},{linha}]", False) is not None:
            nomeJob = session.findById(
                f"wnd[0]/usr/lbl[{pos_nome_job},{linha}]").text
            statusJob = session.findById(
                f"wnd[0]/usr/lbl[{pos_status},{linha}]").text
            jobConcluido = statusJob == "Concl."  # RETORNA ZERO SE AS STRINGS FOREM IGUAIS
            # SE NAO É CABECALHO OU RODAPE, PEGAR A HORA PROG DO JOB
            if nomeJob != "NomeJob" and nomeJob != "Resumo" and jobConcluido:
                horaProg = time.strptime(session.findById(
                    f"wnd[0]/usr/lbl[{pos_hora_inicio},{linha}]").text, "%H:%M:%S")

        for dict_job in dados_json:

            if type(dict_job["path"]) == list:
                for job_path in dict_job["path"]:
                    caminho = job_path
                    salvar_arquivo(dict_job, nomeJob,
                                   horaProg, jobConcluido, linha, caminho)
            else:
                caminho = dict_job["path"]
                salvar_arquivo(dict_job, nomeJob,
                               horaProg, jobConcluido, linha, caminho)

        jobConcluido = False
    return


# Salva arquivo
def salvar_arquivo(dict_job, nomeJob, horaProg, jobConcluido, linha, caminho):
    global SapGuiAuto
    global application
    global connection
    global session

    session.findById("wnd[0]").maximize()
    try:
        if dict_job["ativo"] and nomeJob == dict_job["nome_job"] and horaProg == time.strptime(dict_job["hora_inicio"], "%H:%M:%S") and jobConcluido:
            # DELETAR ARQUIVO existente
            try:
                if os.path.exists(caminho + "\\" + dict_job["nome_arquivo"]):
                    os.remove(caminho +
                              "\\" + dict_job["nome_arquivo"])
            except:
                pass

            session.findById(
                f"wnd[0]/usr/chk[1,{linha}]").selected = -1
            session.findById("wnd[0]/tbar[1]/btn[44]").press()
            # Se for transação de material, nome "RIAUFMVK", tem um passo a mais
            if nomeJob == "RIAUFMVK":
                session.findById("wnd[0]/usr/lbl[1,3]").SetFocus()
                session.findById(
                    "wnd[0]/usr/lbl[1,3]").caretPosition = 0
                session.findById("wnd[0]/tbar[1]/btn[34]").press()
            if session.findById("wnd[0]/sbar").text != "Não há lista":
                session.findById("wnd[0]/usr/chk[1,3]").selected = -1
                session.findById("wnd[0]/usr/chk[1,3]").setFocus()
                session.findById("wnd[0]/tbar[1]/btn[6]").press()
                session.findById("wnd[0]/tbar[1]/btn[48]").press()
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                session.findById(
                    "wnd[1]/usr/ctxtDY_PATH").text = caminho
                session.findById(
                    "wnd[1]/usr/ctxtDY_FILENAME").text = dict_job["nome_arquivo"]
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            session.findById(f"wnd[0]/usr/chk[1,{linha}]").selected = 0
            jobConcluido = False
    except:
        pass

    return


# Configura novos relatório background para a data e horário definidos
def gerar_novo_background(dados_json):
    global SapGuiAuto
    global application
    global connection
    global session

    #SapGuiAuto = win32com.client.GetObject('SAPGUI')
    #application = SapGuiAuto.GetScriptingEngine
    #connection = application.Children(0)
    #session = connection.Children(0)

    session.findById("wnd[0]").maximize()

    for dict_job in dados_json:
        if dict_job["ativo"] and "programar_proximo_background" in dict_job:
            if dict_job["programar_proximo_background"]:
                session.findById(
                    "wnd[0]/tbar[0]/okcd").text = f"/n{dict_job['transacao']}"
                session.findById("wnd[0]").sendVKey(0)
                if "variante" in dict_job and len(dict_job["variante"]) > 0:
                    session.findById("wnd[0]/tbar[1]/btn[17]").press()
                    session.findById(
                        "wnd[1]/usr/txtV-LOW").text = f"{dict_job['variante']}"
                    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
                    session.findById("wnd[1]/tbar[0]/btn[8]").press()
                # PROGRAMA NOVO BACKGROUND
                session.findById("wnd[0]/mbar/menu[0]/menu[2]").select()
                session.findById("wnd[1]/tbar[0]/btn[13]").press()
                session.findById("wnd[1]/usr/btnDATE_PUSH").press()
                if "data_inicio_programado" in dict_job:
                    data_inicio_programado = (datetime.now(
                    ) + timedelta(days=dict_job["data_inicio_programado"])).strftime("%d.%m.%Y")
                else:
                    data_inicio_programado = datetime.now().strftime("%d.%m.%Y")
                session.findById(
                    "wnd[1]/usr/ctxtBTCH1010-SDLSTRTDT").text = data_inicio_programado
                session.findById(
                    "wnd[1]/usr/ctxtBTCH1010-SDLSTRTTM").text = dict_job["hora_inicio_programado"]
                session.findById(
                    "wnd[1]/usr/ctxtBTCH1010-SDLSTRTTM").setFocus()
                session.findById(
                    "wnd[1]/usr/ctxtBTCH1010-SDLSTRTTM").caretPosition = 8
                session.findById("wnd[1]/tbar[0]/btn[11]").press()
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
    return


# Cria arquivo CondCargaPIM.txt
def criar_arquivo_controle(dados_json):
    try:
        # Pega o primeiro registro json, pra determinar o diretório onde será salvo o arquivo.
        os.remove(dados_json[0]["path"] + "\\ArquivosCond\\CondCargaPIM.txt")
    except:
        pass
    arquivo = open(dados_json[0]["path"] +
                   "\\ArquivosCond\\CondCargaPIM.txt", "w")
    arquivo.close()
    return


# Chama as funções do SAP
def executar_script_sap(nome_arquivo_json=""):
    # Ler o arquivo JSON com os parâmetros
    arquivo_json = nome_arquivo_json
    #arquivo_json = "arquivos_pim_confirmacoes_14h.json"
    with open(arquivo_json, encoding='utf-8') as f:
        dados = json.load(f)

    salvar_arquivo_background(dados)
    gerar_novo_background(dados)
    criar_arquivo_controle(dados)

    return


if __name__ == "__main__":
    sap_login()  # ---------------------------------- passo 1
    if len(sys.argv) == 1:
        nome_arquivo_json = "arquivos_pim_geral.json"
    else:
        nome_arquivo_json = sys.argv[1]
    executar_script_sap(nome_arquivo_json)  # ------- passo 2
    sap_logoff()  # --------------------------------- passo 3

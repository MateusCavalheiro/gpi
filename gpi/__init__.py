from sap_module import Sap_automato
from pim_module import Q_pim
from gpi_module import Gpi
import pandas as pd
from sap_module import Sap_automato
from datetime import datetime, timedelta
import pyperclip
import time
import win32com.client
import os

rpa = Gpi() #cria o objeto rpa da classe GPI

semana = datetime.today().isocalendar().week + 1  ## Configuração para pegar a semana posterior (atual+1)
ano = datetime.today().isocalendar().year # Ano atual

# Defina o caminho da pasta de destino
pasta_destino = f"C:/Users/{os.getlogin()}/PETROBRAS/GPIs Refino - 08.  Programação Semanal" 

Chamada_rpa = rpa.save_schedule_gpisrefino(semana,ano, pasta_destino, "PIM-PROG-SEDE", "REPAR")


"""
print(teste)

def salvar_ultimo_arquivo_excel(caminho_destino, nome_planilha, ano, semana):
        
    try:
        # Conectar ao Excel se estiver aberto
        excel = win32com.client.GetActiveObject("Excel.Application")

        # Verificar se há arquivos abertos
        if excel.Workbooks.Count == 0:
            print("Nenhum arquivo do Excel está aberto.")
            return
        
        # Selecionar o último arquivo aberto
        ultimo_arquivo = excel.Workbooks(excel.Workbooks.Count)

        # Definir o caminho de destino
        if nome_planilha == None:
            nome_arquivo = os.path.basename(ultimo_arquivo.FullName + '.xlsx')  # Pega o nome original do arquivo
        else:
            nome_arquivo = nome_planilha + '_' +str(ano) + '_' + str(semana) + '.xlsx'
        caminho_completo = os.path.join(caminho_destino, nome_arquivo)

        # Salvar cópia na nova pasta
        ultimo_arquivo.SaveCopyAs(caminho_completo)

        print(f"Arquivo salvo em: {caminho_completo}")

    except Exception as e:
        print(f"Erro ao salvar o arquivo: {e}")

#query_pim = Q_pim()

#obtem pasta 
#df = query_pim.q_programacao()

# Remove duplicadas da coluna e copia para o clipboard CENTRO DE TRABALHO
#df_sem_duplicatas_centroTrabalho = df.drop_duplicates(subset=["centroTrabalho"])


# Remove duplicadas da coluna e copia para o clipboard ORDEM
#df_sem_duplicatas_ordem = df.drop_duplicates(subset=["ordem"])

# Obter o número da semana
#numero_semana = datetime.today().isocalendar()

segunda_feira = datetime.fromisocalendar(datetime.today().isocalendar().year,(datetime.today().isocalendar().week + 1), 1)
sexta_feira = segunda_feira + timedelta(days=4)
segunda_sap =  segunda_feira.strftime('%d.%m.%Y')
sexta_sap =  sexta_feira.strftime('%d.%m.%Y')

session = Sap_automato().sap_login()
#session = Sap_automato().abrir_sap()

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").text = "/niw37"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]").sendVKey(17)
session.findById("wnd[1]/usr/txtV-LOW").text = "PIM-PROG-SEDE"
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[1]").sendVKey(8) #f8 para confirmar
###nesta etapa fazer clipboard dos centros
#df_sem_duplicatas_centroTrabalho['centroTrabalho'].to_clipboard(index=False, header=False)
session.findById("wnd[0]/usr/btn%_ARBPL_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]").sendVKey(16) #Apaga todos os dados
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "*" #Insere o asterisco para pegar todos valores
session.findById("wnd[0]").sendVKey(8) #executa com o clipboard atual
###nesta etapa fazer clipboard das ordens
#df_sem_duplicatas_ordem['ordem'].to_clipboard(index=False, header=False)
session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press() #abrir para inserir ordem
session.findById("wnd[1]").sendVKey(16) #Apaga todos os dados
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "*"
session.findById("wnd[1]").sendVKey(8) #f8 para confirmar
session.findById("wnd[0]/usr/ctxtFSAVD-LOW").text = segunda_sap #insere a data de início (segunda feira)
session.findById("wnd[0]/usr/ctxtFSAVD-HIGH").text = sexta_sap #insere a data de fim (sexta feira)
session.findById("wnd[0]").sendVKey(8) #executa a consulta
session.findById("wnd[0]").sendVKey(16) #alt+f4 para gerar planilha
session.findById("wnd[1]").sendVKey(0)
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus()
time.sleep(2)
session.findById("wnd[1]").sendVKey(0)
time.sleep(2)
session.findById("wnd[1]").sendVKey(0)
time.sleep(2)


# Defina o caminho da pasta de destino
pasta_destino = f"C:/Users/{os.getlogin()}/PETROBRAS"#/GPIs Refino - 08.  Programação Semanal" 
#pasta_destino = f"C:/Users/{os.getlogin()}/PETROBRAS/Controle da Manutenção MA - EI - Imagens Válvulas e Dampers"

# Criar a pasta se não existir
os.makedirs(pasta_destino, exist_ok=True)

# Salvar o último arquivo do Excel na pasta escolhida
salvar_ultimo_arquivo_excel(pasta_destino, 'REPAR', datetime.today().isocalendar().year, datetime.today().isocalendar().week + 1)

session.findById("wnd[0]").close()
session.findById("wnd[2]/usr/btnSPOP-OPTION1").press()

#Sap_automato.sap_logoff()"""
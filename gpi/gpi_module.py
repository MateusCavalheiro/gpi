import pyautogui, sys, subprocess
import time as tm
import win32com.client
import os
from pandas import DataFrame
from sap_module import Sap_automato
from general_module import Action_excel
from datetime import datetime, timedelta
from yaspin import yaspin

class Gpi:
    def __init__(self, value=1):
        self.value = value
        
    def save_schedule_gpisrefino(self, semana, ano, pasta_destino, variante, nome_planilha) -> str:
        """RPA para salvamento da programação no formato definido pela SEDE sem filtro de centros ou ordens das programadas controladas.

        Returns:
            str: [Local / data / hora / semana] do arquivo salvo
        """
        


        with yaspin(text="Carregando...", color="cyan") as spinner:
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
            
            segunda_feira = datetime.fromisocalendar(ano,semana, 1)
            sexta_feira = segunda_feira + timedelta(days=4)
            segunda_sap =  segunda_feira.strftime('%d.%m.%Y')
            print("Segunda feira utilizada: " + segunda_sap)
            sexta_sap =  sexta_feira.strftime('%d.%m.%Y')
            print("Sexta feira utilizada: " + sexta_sap)
            
            session = Sap_automato().sap_login()
            
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "/niw37"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").sendVKey(17)
            session.findById("wnd[1]/usr/txtV-LOW").text = variante #"PIM-PROG-SEDE"
            session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(8) #f8 para confirmar
            
            session.findById("wnd[0]/usr/btn%_ARBPL_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]").sendVKey(16) #Apaga todos os dados
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "*" #Insere o asterisco para pegar todos valores
            session.findById("wnd[0]").sendVKey(8) #executa com o clipboard atual
            
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
            tm.sleep(2)
            session.findById("wnd[1]").sendVKey(0)
            tm.sleep(2)
            session.findById("wnd[1]").sendVKey(0)
            tm.sleep(2)
            
            # Criar a pasta se não existir
            os.makedirs(pasta_destino, exist_ok=True)
            # Salvar o último arquivo do Excel na pasta escolhida
            salvar_ultimo_arquivo_excel(pasta_destino, nome_planilha, ano, semana)
            
            spinner.ok("✅ Concluído!")
            return segunda_sap + " / " + sexta_sap
        
    
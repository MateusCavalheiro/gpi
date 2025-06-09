import pyautogui, sys, subprocess
import time as tm
import win32com.client
import os
from pandas import DataFrame
from sap_module import Sap_automato
from general_module import Action_excel
from datetime import datetime, timedelta
from yaspin import yaspin
import pyperclip
import pandas as pd

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
        
    def controle_materiais(self, Perfil_MRP, Centro, Planejador: list, Roll_start: int=0, tentativas: int=0, Roll_end: int=0) -> DataFrame:
        """RPA para controle de materiais.

        Args:
            Perfil_MRP: Perfil MRP.
            Centro: Centro.
            Planejador (list): Lista de planejadores.

        Returns:
            DataFrame: DataFrame com os dados do controle de materiais.
        """
        with yaspin(text="Carregando...", color="cyan") as spinner:
            session = Sap_automato().sap_login()
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nMMD7" # Escreve transação
            session.findById("wnd[0]").sendVKey(0) # Enter para entrar
            
            tm.sleep(1)
            
            session.findById("wnd[0]/usr/ctxtDISPRO").text = Perfil_MRP # Coloca o tipo de perfil MRP
            session.findById("wnd[0]/usr/ctxtWERK-LOW").text = Centro # Centro de trabalho
            session.findById("wnd[0]/usr/btn%_DISPON_%_APP_%-VALU_PUSH").press() # ABRE PARA INSERIR MAIS DE UM
            session.findById("wnd[1]").sendVKey(16) # Garante que não tera nenhum valor antes de colar

            # Copiar a lista de planejadores para o clipboard usando pandas
            pd.Series(Planejador).to_clipboard(index=False, header=False)
            session.findById("wnd[1]").sendVKey(24) # Cola o Clipboard que é feito com a lista passada no parametro (Não pode ser vazia)
            session.findById("wnd[1]").sendVKey(8) # confirma pesquisa
            session.findById("wnd[0]").sendVKey(8) #Entra com a pesquisa feita
            
            # Identificar a grade
                #session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(-1, "MTART")
                #session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn("MATNR")
                #session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn("MAKTX")
                #session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn("DISPO")
                #session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn("MTART")
                #session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = ""
            
            #session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "10"
            #session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").copySelectedRowsToClipboard


            


            # Assume que `session` já está corretamente inicializado
            grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
            
            if Roll_end > 0:
                lines_grid = Roll_end
            else:
                lines_grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").RowCount
            print(lines_grid) # Obtem o número de linhas da grade
            #session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleRow = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").RowCount -35
            last_roll = lines_grid - 35
            
            i = Roll_start
            count_errors = 0
            print('Selecione o SAP')
            for tim in range(10):
                print(f"Esperando {10-tim} segundos para selecionar o SAP...")
                tm.sleep(1)
            spinner.text = "Copiando dados..."
            if Roll_start > 0:
                session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleRow = Roll_start  # Define a linha inicial para rolagem
            while i < lines_grid+1:
                try:
                    grid.selectedRows = str(i)  # Seleciona a linha atual
                    
                    tm.sleep(0.3)  # Aguarda seleção
                    pyautogui.hotkey('ctrl', 'c')  # Ctrl + C para copiar os dados da grade
                    clipboard_data = pyperclip.paste()
                    print(clipboard_data)
                    # Lê os dados copiados da área de transferência em um DataFrame temporário
                    #temp_df = pd.read_clipboard()
                    temp_df = pd.read_clipboard(sep='\t', names=["Centro", "Área MRP", "Material", "Texto Breve Material", "P|MRP", "TMat"], header=None)
                
                    # Concatena ao DataFrame principal
                    if i == Roll_start:
                        df = temp_df
                    else:
                        df = pd.concat([df, temp_df], ignore_index=True)
                        
                    if i >= last_roll:
                        i +=1
                        
                    else: 
                        i += 1
                        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleRow += 1
                    print(f"Dados da linha {i} copiados com sucesso.")
                except Exception as e:
                    print(f"Erro ao copiar dados da linha {i}: {e}") 
                    i+=1
                    count_errors += 1
                    if count_errors > 10:
                        spinner.fail("❌ Muitos erros ao copiar dados.")
                        tm.sleep(2)
                        print("Muitos erros ao copiar dados. Encerrando o processo.")
                        if tentativas <= 5:
                            print(f"Tentativa {tentativas+1} de 5 para reiniciar o processo.")
                            print("Reiniciando o processo...")
                            print("Com os dados: ",Perfil_MRP, Centro, Planejador, i-10, tentativas+1)
                            temp_df = self.controle_materiais(Perfil_MRP, Centro, Planejador, i-10, tentativas + 1)
                            df = pd.concat([df, temp_df], ignore_index=True)
                        else:
                            break
                
            #pyautogui.hotkey('ctrl', 'c')  # Ctrl + C para copiar os dados da grade
            
            
            
            spinner.ok("✅ Concluído!")
            
            
            return df

if __name__ == "__main__":
    
    #Cria o objeto da rpa
    rpa = Gpi()

    Chamada_rpa = rpa.controle_materiais(
        Perfil_MRP="ZDZM",
        Centro="1400",
        Planejador=["ME1", "MI1"],
        Roll_start=0,
        Roll_end = 1000
    )

    # Carrega o DataFrame existente, se o arquivo já existir
    output_file = "Material com perfil MRP ZDZM ZD ZM.xlsx"
    if os.path.exists(output_file):
        df_existente = pd.read_excel(output_file)
        df_final = pd.concat([df_existente, Chamada_rpa], ignore_index=True)
    else:
        df_final = Chamada_rpa

    df_final.to_excel(output_file, index=False)
    print(f"Dados salvos no arquivo: {output_file}")
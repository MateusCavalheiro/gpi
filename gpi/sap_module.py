import pyautogui, sys, subprocess

import time as tm
import win32com.client



class Sap_automato:
    def __init__(self, value=1):
        self.value = value
            
            
    def chamada_sap(self, s48=False):
        global SapGuiAuto
        global application
        global connection
        global session
        #
        
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
            aba = connection.children.Count
            
            print("Quantidade de janelas", aba)
            if aba > 0 and aba < 6:
                session.createSession()
                tm.sleep(2)
                session = connection.Children(aba)            
                session.findById("wnd[0]").maximize()
            else:
                print("Será utilizada a última aba aberta")
                tm.sleep(2)
                
                session = connection.Children(aba-1)            
                session.findById("wnd[0]").maximize()
            return aba    
        except:
            try:
                path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
                subprocess.Popen(path)
                tm.sleep(5)   

            except:

                window_list = pyautogui.getWindowsWithTitle("SAP Logon 770")

                if window_list:
                # Suponhamos que você queira reativar a primeira janela na lista
                    window_to_activate = window_list[0]

                # Ativar a janela
                    window_to_activate.activate()

                else:
                    print("Nenhuma janela encontrada com o título especificado.")

            try:    
                SapGuiAuto = win32com.client.GetObject('SAPGUI')
                application = SapGuiAuto.GetScriptingEngine
                if s48:
                    connection = application.OpenConnection(
                    r"S48 [SAPSCRIPT]", True)
                else:    
                    connection = application.OpenConnection(
                    r"02 PEP - SAP S/4HANA Produção (SAP SCRIPT)", True)
                
                session = connection.Children(0)
                session.findById("wnd[0]").maximize()
                print('Sap Aberto............')
                return aba
            except Exception as e:
                print("Não encontrado o 02 PEP - SAP S/4HANA Produção (SAP SCRIPT)")
                print(e)
            
            
    def abrir_sap(self):
        """
        This function just open sap in your computer.
        
        Esta função apenas abre o sap em seu computador.
        """
        
        janela = self.chamada_sap()
        
        
        janela_count_add = janela + 1 if janela <6 else 6
        
        
        print(f'SAP aberto com sucesso na janela:{janela_count_add}')
        

    def sap_login(self) -> object:
        """Função para realizar abertura do sap

        Returns:
            session: Retorna a sessão para interagir com SAP
        """
        
        
        global SapGuiAuto
        global application
        global connection
        global session

        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
            aba = connection.children.Count
            
            print("Quantidade de janelas", aba)
            if aba > 0 and aba < 6:
                session.createSession()
                tm.sleep(2)
                session = connection.Children(aba)            
                session.findById("wnd[0]").maximize()
            else:
                print("Será utilizada a última aba aberta")
                tm.sleep(2)
                
                session = connection.Children(aba-1)            
                session.findById("wnd[0]").maximize()
            return session    
        except Exception as e:
            print("Ocorreu um erro ao tentar abrir uma nova janela na sessão do Sap atual. Abrindo uma nova sessão do SAP.")
            
            try:
                path = r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"
                subprocess.Popen(path)
                tm.sleep(10)
                    
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
                try:
                #Tenta apertar o botão de force para substituir a sessão aberta (é raro mas acontece)
                    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                except:
                    session.findById("wnd[0]").sendVKey(0)
                return session
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
    





#
#def haversine(lon1: float, lat1: float, lon2: float, lat2: float) -> float:
#    """
#    Calculate the great circle distance between two points on the 
#    earth (specified in decimal degrees), returns the distance in
#    meters.
#    All arguments must be of equal length.
#    :param lon1: longitude of first place
#    :param lat1: latitude of first place
#    :param lon2: longitude of second place
#    :param lat2: latitude of second place
#    :return: distance in meters between the two sets of coordinates
#    """
#    
#    
#    # Convert decimal degrees to radians
#    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
#
#    # Haversine formula
#    dlon = lon2 - lon1
#    dlat = lat2 - lat1
#    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
#    c = 2 * asin(sqrt(a))
#    r = 6371  # Raio da Terra em quilômetros
#    return c * r
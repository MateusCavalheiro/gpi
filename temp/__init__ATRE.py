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

semana = datetime.today().isocalendar().week  ## Configuração para pegar a semana atual (atual+1 = posterior)
ano = datetime.today().isocalendar().year # Ano atual

# Defina o caminho da pasta de destino
pasta_destino = f"C:/Users/{os.getlogin()}/PETROBRAS/REPAR GPI - Programação ATREs"

Chamada_rpa = rpa.save_schedule_gpisrefino(semana,ano, pasta_destino, "GPI-01-PROG", "PROG_ATRE")


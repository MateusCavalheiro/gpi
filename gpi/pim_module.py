import pyodbc
import pandas as pd
from datetime import datetime


class Q_pim:
    def __init__(self, value=1):
        self.value = value
    
    def q_programacao(self) -> pd.DataFrame:
        """
        This method will make the connection with PIM and collect schedule of week
        
        *** PARA ACESSAR ESTA BASE DE DADOS VOCÊ PRECISA DO PIM CONFIGURADO NA MÁQUINA EXEMPLO ->[Sujerido consulta no procedimento 0.1](https://petrobrasbr.sharepoint.com/teams/bdoc_REPAR-MA/Documentos%20Compartilhados/Forms/AllItems.aspx?FolderCTID=0x0120005FE5A41C3150C944A5440298FDF866B0&id=%2Fteams%2Fbdoc%5FREPAR%2DMA%2FDocumentos%20Compartilhados%2FGPI%2FINTERNO%2F03%2E%20Procedimentos%20Internos)
        """
        # Caminho para o arquivo .accdb do banco de dados Access
        db_file = r'P:\PIM\bd\tb_PIM.accde'
        
        # Conectar ao banco de dados Access via ODBC
        conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_file)
        # Consulta SQL
        query = """ SELECT DISTINCT
                    VW.ordem
                    , VW.operacao
                    , VW.subOperacao
                    , VW.centroTrabalho
                    , VW.textoBreveOrdem
                    , VW.descricaoOperacao
                    , VW.qtdeExecutante
                    , VW.qtdeHoras
                    , VW.hhPlanejado
                    , VW.dataInicio
                    , VW.horaInicio
                    , VW.dataFim
                    , VW.horaFim
                    , VW.codigoGPM AS GPM
                    , VW.codigoGPMOrdem AS GPMOrdem
                    , VW.areaOperacional
                    , VW.prioridadeOrdem
                    , VW.PFC
                    , VW.PQT
                    , VW.CAP
                    , VW.SGSO
                    , VW.ZF
                    , VW.ZR
                    , VW.ZI
                    , VW.priorizaCriterio
                    , VW.localizacao
                    , VW.AR
                    , VW.LIBRA
                    , VW.ARO
                    FROM vw_programacaoSemanal AS VW
                    WHERE VW.anoSemana IN (SELECT DISTINCT ads_SEMANA FROM tbWork_ANALISE_DETALHE_SEMANAL WHERE ads_DT_CARGA = (SELECT MAX(ads_DT_CARGA) FROM tbWork_ANALISE_DETALHE_SEMANAL))
                    ORDER BY
                    VW.areaOperacional
                    , VW.ordem
                    , VW.dataInicio
                    , VW.horaInicio"""
        # Executando a consulta
        df = pd.read_sql(query, conn)
        # Fechar a conexão
        conn.close()
        
        return pd.DataFrame(df)



        
        
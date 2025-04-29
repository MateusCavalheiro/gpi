import win32com.client, os



class Action_excel:
    @staticmethod
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
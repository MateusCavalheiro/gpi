import pyodbc
import os, json

arquivo_json = os.path.join(os.environ['USERPROFILE'], "config_robos.json")

with open(arquivo_json, encoding='utf-8') as f:
    dados = json.load(f)

usuarioID = dados["senha_rede"]["chave"]
senhaAcesso = dados["senha_rede"]["senha"]

pasta_arquivo = dados["SPT"]["destino_exportacao"]
nome_arquivo = dados["SPT"]["nome_arquivo"]

# Listar todos os drivers ODBC disponíveis
drivers = pyodbc.drivers()
print("Drivers ODBC instalados:")
for driver in drivers:
    print(f"- {driver}")

# Listar DSNs (Data Source Names) configurados
dsns = pyodbc.dataSources()
print("\nDSNs configurados:")
for dsn in dsns:
    print(f"- {dsn}: {dsns[dsn]}")
    
    
    
    """def df_ods_query(query, generate_plan, name_plan):
    start_time = datetime.datetime.now()
    dsn_tns = oracledb.makedsn(
        'bdodscp.petrobras.com.br',
        '1521',
        service_name='odscp.petrobras.com.br'
    )

    try:
        with oracledb.connect(user='d25e', password='b#S03mkrdf', dsn=dsn_tns) as connection: # SENHA DO SQLDEVELOPER
            with connection.cursor() as cursor:
                cursor.execute(query)
                
                
                
                df = pd.read_sql_query(query, connection)
    except oracledb.Error as error:
        print(error)
    
    if generate_plan:
        cont= 0
        #print(cont)
        while True:
            try:
                if cont == 0:
                    df.to_excel(name_plan + '.xlsx')
                else:
                    df.to_excel(name_plan + str(cont) +'.xlsx')
                break
            except:
                cont += 1
    end_time = datetime.datetime.now()
    time_exec = (end_time - start_time).seconds
    return df, time_exec
    """
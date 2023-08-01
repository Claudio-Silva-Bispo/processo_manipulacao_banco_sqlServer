#--------------------( Criar função de conexão )--------------------#

import pymssql

def conectar_sql():
    server = 'tms-drp-oficial.database.windows.net'
    database = 'tms'
    user = 'appTableau'
    password = 'Tjd82#jks2'

    conn = pymssql.connect(server=server, database=database, user=user, password=password)

    return conn


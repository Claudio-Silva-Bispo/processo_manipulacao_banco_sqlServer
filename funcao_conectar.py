#--------------------( Criar função de conexão )--------------------#

import pymssql

def conectar_sql():
    server = ''
    database = ''
    user = ''
    password = ''

    conn = pymssql.connect(server=server, database=database, user=user, password=password)

    return conn


from datetime import datetime

import pyodbc
import pandas as pd


def connection(db: object, ip_server: object, user: object, pwd: object) -> object:
    connect = None
    try:
        connect = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};'
            f'SERVER={ip_server};'
            'PORT=1433;'
            f'DATABASE={db};'
            f'UID={user};'
            f'PWD={pwd}'
        )
        return connect
    except pyodbc.Error as err:
        print(f'Error: {str(err)}')
        input('Aperte Enter para sair...')
        quit()


def consulta(connect: object, query: object) -> object:
    cursor = connect.cursor()
    result = None
    try:
        cursor.execute(query)
        result = cursor.fetchall()
        cursor.close()
        return result
    except cursor.Error as err:
        print(f'Error : {str(err)}')
        input('Aperte Enter para sair...')
        quit()



def list_db(results):
    from_db = []
    for result in results:
        result = list(result)
        from_db.append(result)

    col = ['CODIGO', ' NF-O  ', ' QTATD  ', ' DTEMS  ', ' PRUNI ', ' ICMS  ', ' PCIPI', ' CFOP', ' RAZÃO ', ' CNPJ   ',
           ' CHAVE_ACESSO ', ' CDPRD ', ' CDFRN ', ' UFFRN ', ' DTENT']
    df = pd.DataFrame(from_db, columns=col)
    return df


def import_excel(data, cod, loj):
    data.to_excel(r'Prod '+ str(cod) +' Loja '+str(loj)+' - ' + str(datetime.now().strftime("%H.%M.%S-%Y")) + '.xlsx', index=False)
    return print('Saida em Excel')



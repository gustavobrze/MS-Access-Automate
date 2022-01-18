import pyodbc

def new_access(sapid, name, copy_id, table):

    with pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=C:\Users\OEM\Downloads\scripts\accessautomate\db1.accdb;') as conn: # Cria a conexão com a base de dados

        cursor = conn.cursor()  # Cria o cursor

        cursor.execute('SELECT * FROM ' + table + ' WHERE usuario = ?', copy_id) # Seleciona as linhas do usuário espelho
        copy_user = cursor.fetchall()

        for i in copy_user:

            new_access = list()
            chave_user = str(sapid) + "-" + i[2]  # Cria a Chave com o id do usuário

            cursor.execute('SELECT COUNT(Chave) FROM ' + table + ' WHERE Chave = ?', chave_user) # Verifica valores duplicados na Chave
            row = cursor.fetchone()
            
            if row[0] <= 0:   # Se <= 0, a chave não foi identificada, portanto, o acesso é adicionado
                new_access.append((sapid, name, i[2], i[3], i[4],
            i[5], i[6], str(sapid) + "-" + i[2], i[8], i[9]))
                cursor.executemany('INSERT INTO ' + table + ' VALUES(?,?,?,?,?,?,?,?,?,?)', (new_access))

        cursor.commit()

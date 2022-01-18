import pyodbc

def new_access(sapid, name, copy_id):

    with pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=C:\Users\OEM\Downloads\db1.accdb;') as conn: # Cria a conexão com a base de dados

        cursor = conn.cursor()  # Cria o cursor

        cursor.execute('SELECT * FROM UsuarioEquipamentos WHERE usuario = ?',copy_id)
        copy_user = cursor.fetchall()

        for i in copy_user:

            new_access = list()
            chave_user = str(sapid) + "-" + i[2]

            cursor.execute('SELECT COUNT(Chave) FROM UsuarioEquipamentos WHERE Chave = ?',chave_user)
            row = cursor.fetchone()
            if row[0] <= 0:
                new_access.append((sapid, name, i[2], i[3], i[4],
            i[5], i[6], str(sapid) + "-" + i[2], i[8], i[9]))
                cursor.executemany('INSERT INTO UsuarioEquipamentos VALUES(?,?,?,?,?,?,?,?,?,?)',(new_access))

        cursor.commit()

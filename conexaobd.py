from email import charset

import firebirdsql

conn = firebirdsql.connect(
    user="SYSDBA", 
    password="masterkey",
    database="C:\\SAFIBDV3\\SAFI_D3F2_TESTE.FDB",
    host="LOCALHOST",
    charset="ANSI",
    )


#cur = conn.cursor()
#cur.execute("select * from aliq_icms")
#
#for c in cur.fetchall():
#    print(c)

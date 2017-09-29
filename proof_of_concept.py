import pyodbc


print [x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\RSVP_Export_Process\RSVP_Backup_20170928.mdb;'
    )

connection = pyodbc.connect(conn_str)
cursor = connection.cursor()

for item in cursor.execute("SELECT TOP 10 * FROM tblRSVPCases"):
    print item
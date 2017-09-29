import pyodbc
from docx import Document

# print [x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\RSVP_Export_Process\RSVP_Backup_20170928.mdb;'
    )

connection = pyodbc.connect(conn_str)
cursor = connection.cursor()

for item in cursor.execute("SELECT TOP 10 * FROM tblRSVPCases"):
    output_doc = Document()

    case_name = str(item[5]).replace(r'/', '-')

    output_doc.save(r"output/" + case_name + '.docx')



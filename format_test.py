from docx import Document
from docx.enum.section import WD_ORIENTATION

NUM_COLUMNS = 5
TEST_ROLES = ['NCP', 'CP', 'NCP Attorney', 'CP Attorney', 'Child 1', 'Child Attorney']
TEST_DATA = ['Apple Jacks', '123-456-7890', 'apple@jacks.com', 'some notes']

document = Document()
section = document.sections[-1]
section.orientation = WD_ORIENTATION.LANDSCAPE

table = document.add_table(cols=NUM_COLUMNS, rows=1) #just declaring header row
table.style = 'TableGrid'
header_cells = table.rows[0].cells
header_cells[0].text = 'ROLE'
header_cells[1].text = 'NAME'
header_cells[2].text = 'PHONE #s'
header_cells[3].text = 'EMAIL'
header_cells[4].text = 'NOTES'

for role in TEST_ROLES:
    cells = table.add_row().cells
    cells[0].text = role
    for i in xrange(0, len(TEST_DATA)):
        cells[i+1].text = TEST_DATA[i]

document.save(r'output/test.docx')




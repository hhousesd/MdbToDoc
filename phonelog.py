from docx import Document
from docx.enum.section import WD_ORIENTATION

NUM_COLUMNS = 5
TEST_ROLES = ['NCP', 'CP', 'NCP Attorney', 'CP Attorney', 'Child 1', 'Child Attorney']
TEST_DATA = ['Apple Jacks', '123-456-7890', 'apple@jacks.com', 'some notes']
OUTPUT_DIR = r'output/'
OUTPUT_EXT = '.docx'

class PhoneLog:

    def __init__(self, case_name):
        self.case_name = case_name
        self.document = Document()
        section = self.document.sections[-1]
        section.orientation = WD_ORIENTATION.LANDSCAPE

    def add_people(self, people):
        table = self.document.add_table(cols=NUM_COLUMNS, rows=1) #just declaring header row
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

    def save(self):
        self.document.save(OUTPUT_DIR + self.case_name + OUTPUT_EXT)

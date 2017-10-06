from docx import Document
from docx.enum.section import WD_ORIENTATION

NUM_COLUMNS = 5
OUTPUT_DIR = r'output/'
OUTPUT_EXT = '.docx'
ROLES_WITH_ADDRESS = ['cp', 'ncp', 'cp attorney', 'ncp attorney']


class PhoneLog:
    def __init__(self, case_name):
        self.case_name = case_name
        self.document = Document()
        section = self.document.sections[-1]
        section.orientation = WD_ORIENTATION.LANDSCAPE

    def add_people(self, people):
        table = self.document.add_table(cols=NUM_COLUMNS, rows=1)  # just declaring header row
        table.style = 'TableGrid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'ROLE'
        header_cells[1].text = 'NAME'
        header_cells[2].text = 'PHONE #s'
        header_cells[3].text = 'EMAIL'
        header_cells[4].text = 'NOTES'

        # TODO need to handle multiple phone numbers per Person from Table:tblPhoneNumbers

        for person in people:
            role = str(person[2])
            first_name = person[3]
            last_name = person[4]
            sex = person[6]
            address_street = person[7]
            address_city = person[8]
            address_state = person[9]
            address_postal = person[10]
            address_country = person[11]
            notes = person[12]

            if first_name is None or str(first_name).strip().isspace() or \
                            last_name is None or str(last_name).strip().isspace():
                continue

            cells = table.add_row().cells
            cells[0].text = role
            cells[1].text = str.format("{0} {1}", first_name, last_name)
            cells[2].text = "placeholder phonenumber"

            role_normalized = role.strip().lower()
            if role_normalized == 'child':
                cells[3].text = str(sex)
            else:
                cells[3].text = ""

            if role_normalized in ROLES_WITH_ADDRESS and address_street is not None:
                cells[4].text = \
                    str.format("Address: {street}, {city}, {state}, {zip}, {country}",
                               street=address_street, city=address_city, state=address_state,
                               zip=address_postal, country=address_country)

    def save(self):
        self.document.save(OUTPUT_DIR + self.case_name + OUTPUT_EXT)

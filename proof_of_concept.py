import pyodbc
from docx import Document
from docx.enum.section import WD_ORIENTATION

from phonelog import PhoneLog
from queries.qryTestGroupPeopleByCase import QUERY_CASES_AND_PEOPLE

# print [x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]

CONN_STR = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\RSVP_Export_Process\RSVP_Backup_20170928.mdb;'
    )


def get_cases_from_rsvp_mdb():
    connection = pyodbc.connect(CONN_STR)
    cursor = connection.cursor()
    results = {}

    for item in cursor.execute(QUERY_CASES_AND_PEOPLE):

        case_id = item[0]

        if not results.has_key(case_id):
            results[case_id] = []
        results[case_id].append(item)

    return results


def format_output_filename(case_name):
    return case_name.strip().replace(r'/', '-')


cases = get_cases_from_rsvp_mdb()
for case_id in cases.keys():

    if len(cases[case_id]) == 0:
        print(str.format("case id {0} has no associated people", case_id))

    case_name = str(cases[case_id][0][1])
    if not case_name.isspace():
        case_name = format_output_filename(case_name)

    people = cases[case_id]
    phone_log = PhoneLog(case_name)
    phone_log.add_people(people)
    phone_log.save()

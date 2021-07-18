from .parse_xls import process_field
import settings
from docx import Document

ROW_ID = 0


def parse_data(docx):
    docx_rows = {}
    docx_rows_generator = (row for row in docx.tables[0].rows[1:] if row.cells[1].text)
    global ROW_ID
    for row in docx_rows_generator:
        dict_row = {}
        for i in settings.DOCX_COLUMN_INDEXES:
            value = row.cells[i].text
            dict_row[settings.DOCX_COLUMNS[i]['var']] = process_field(value,
                                                                      settings.DOCX_COLUMNS[i]['type'])
        docx_rows[ROW_ID] = dict_row
        ROW_ID += 1
    return docx_rows


def parse(filename):
    document = Document(filename)
    data = parse_data(document)
    return data

import locale
from datetime import datetime
import re
import settings
import pandas as pd


ROW_ID = 0


def process_field(field, field_type):
    if field_type in ('int', 'float'):
        return str(field)
    elif field_type == 'date':
        return process_date(field)
    return field.strip()


def process_date(date):
    if not date or isinstance(date, datetime):
        return date
    if len(date) < 8:
        return None
    locale.setlocale(locale.LC_TIME, 'ru')
    date = date.lower().strip()
    groups = [*re.match(r'(\d+).*?([А-Яа-я0-9]+).*?(\d+)', date).groups()]

    if len(groups[2]) < 4:
        year_format = 'y'
    else:
        year_format = 'Y'
    if groups[1].isalpha():
        groups[1] = settings.MONTH_DECLENSIONS.get(groups[1], groups[1])
        month_format = 'B'
    else:
        month_format = 'm'
    date = '.'.join(groups)
    pattern = f'%d.%{month_format}.%{year_format}'
    date = datetime.strptime(date, pattern)
    return date


def parse(filename):
    rows = {}
    df = pd.read_excel(filename)
    df = df.fillna('')
    global ROW_ID
    for row in df.values:
        dict_row = {}
        for i in settings.XLSX_COLUMN_INDEXES:
            value = row[i]
            dict_row[settings.XLSX_COLUMNS[i]['var']] = process_field(row[i],
                                                                      settings.XLSX_COLUMNS[i]['type'])
        rows[ROW_ID] = dict_row
        ROW_ID += 1
    return rows

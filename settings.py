from datetime import datetime
# Словарь для отображения типа совпадения в имя листа excel:
#     'Тип совпадения': 'Имя в выходном файле'
MATCH_TYPE_TO_SHEET_LABEL = {'full': 'Полные совпадения', 'name_num': 'Совпадения по названию и номеру',
                             'name_date': 'Совпадения по названию и дате',
                             'num_date': 'Совпадения по номеру и дате',
                             'name': 'Совпадения по названию', 'num': 'Совпадения по номеру',
                             'date': 'Совпадения по дате'}

# Название перемменой при отсутствии совпадения
NO_MATCH_VAR = 'no_match'
# Название листа в результирующем xlsx файле
NO_MATCH_SHEET_LABEL = 'Без совпадений'

# Колонка, по которой происходит сортировка, и соответствующее минимальное значение
# date=datetime(1, 1, 1), str='', int=0, float=0
SORT_FIELD = 'date'
MIN_SORT_FILED_VALUE = datetime(1, 1, 1)


# Информация о входных и выходных данных без первого столбца с порядком номером, если он сущестует
# Нумерация с 0 в произвольном порядке
#     Номер столбца: {
#             'var': Псевдоним,
#             'type': Тип столбца',
#             'result_name': Имя колонки в выходном файле
#         }
#     var: Любое уникальное
#     type: 'int, float, str, date'
#     result_name: Любое
DOCX_COLUMNS = {1: {'var': 'num', 'type': 'int', 'result_name': 'Номер договора'},
                2: {'var': 'name', 'type': 'str', 'result_name': 'Наименование Фирмы'},
                3: {'var': 'date', 'type': 'date', 'result_name': 'Дата'},
                4: {'var': 'comment', 'type': 'str', 'result_name': 'Комментарий'}}
XLSX_COLUMNS = {
    1: {'var': 'date', 'type': 'date', 'result_name': 'Дата'},
    2: {'var': 'name', 'type': 'str', 'result_name': 'Наименование Фирмы'},
    3: {'var': 'fio', 'type': 'str', 'result_name': 'ФИО'},
    4: {'var': 'num', 'type': 'int', 'result_name': 'Номер договора'},
    5: {'var': 'comment', 'type': 'str', 'result_name': 'Комментарий'}
}


# Номера колонок из DOCX_COLUMNS и XLSX_COLUMNS в возрастающем порядке
DOCX_COLUMN_INDEXES = (1, 2, 3, 4)
XLSX_COLUMN_INDEXES = (1, 2, 3, 4, 5)


# Номер колонки и ее ширина. По умолчанию длина заголовка * 1.2
DOCX_COLUMNS_WIDTH = {2: 30, 3: 15, 4: 20}
XLSX_COLUMNS_WIDTH = {1: 15, 2: 30, 3: 20, 5: 20}


def match(row1, row2):
    """
    Способ сравнения двух строк в соответствии с DOCX_COLUMNS и XLSX_COLUMNS: row1 - docx, row2 - xlsx
    """
    name_matched = row1['name'] and row2['name'] and row1['name'] == row2['name']
    num_matched = row1['num'] and row2['num'] and row1['num'] == row2['num']
    date_matched = row1['date'] and row2['date'] and row1['date'] == row2['date']
    if name_matched and num_matched and date_matched:
        return 'full'
    elif name_matched and num_matched:
        return 'name_num'
    elif name_matched and date_matched:
        return 'name_date'
    elif num_matched and date_matched:
        return 'num_date'
    elif name_matched:
        return 'name'
    elif num_matched:
        return 'num'
    elif date_matched:
        return 'date'
    else:
        return ''


MATCH_TYPE_TO_SHEET_LABEL.update({NO_MATCH_VAR: NO_MATCH_SHEET_LABEL})

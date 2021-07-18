import xlsxwriter
from settings import (DOCX_COLUMNS, DOCX_COLUMN_INDEXES, DOCX_COLUMNS_WIDTH,
                      XLSX_COLUMNS, XLSX_COLUMN_INDEXES, XLSX_COLUMNS_WIDTH,
                      SORT_FIELD, MIN_SORT_FILED_VALUE, NO_MATCH_VAR,
                      NO_MATCH_SHEET_LABEL, MATCH_TYPE_TO_SHEET_LABEL, match)


def find_row_matches(docx_orders, xlsx_orders):
    matches = dict()
    for i, row1 in docx_orders.items():
        for j, row2 in xlsx_orders.items():
            match_type = match(row1, row2)
            if match_type:
                if match_type not in matches:
                    matches[match_type] = {}
                if i not in matches[match_type]:
                    matches[match_type][i] = []
                matches[match_type][i].append(j)
    return matches


def set_column_name(worksheet, columns, column_indexes, row, col, fmt=None):
    worksheet.write_row(row, col, (columns[i]['result_name'] for i in column_indexes), fmt)


def merge_columns(worksheet, row, first_col, last_col, data, fmt=None):
    worksheet.merge_range(row, first_col, row, last_col, data, fmt)


def set_column_width(worksheet, offset, columns, column_indexes, col_and_width=None):
    for i, col in enumerate(column_indexes, offset):
        length = col_and_width.get(col, len(columns[col]['result_name']) * 1.2)
        worksheet.set_column(i, i, length)


def merge_rows(worksheet, col, first_row, last_row, data, fmt):
    worksheet.merge_range(first_row, col, last_row, col, data, fmt)


def write_row(worksheet, file, file_row, row, col, columns, fmt=None):
    data = [file[file_row][columns[j]['var']] for j in columns.keys()]
    worksheet.write_row(row, col, data, fmt)


def dump_to_excel(matches, docx, xlsx, upload_path='matches.xlsx'):
    n, m = len(DOCX_COLUMNS), len(XLSX_COLUMNS)
    matches[NO_MATCH_VAR] = {}
    remained_xlsx_rows = set(xlsx.keys())
    remained_docx_rows = set(docx.keys())
    with xlsxwriter.Workbook(upload_path) as workbook:
        fmt = workbook.add_format({'num_format': 'dd.mm.yyyy', 'align': 'center', 'border': 2,
                                   'text_wrap': True})
        for match_type in matches:
            worksheet = workbook.add_worksheet(MATCH_TYPE_TO_SHEET_LABEL[match_type])
            merge_columns(worksheet, 0, 0, n - 1, '.DOCX', fmt)
            merge_columns(worksheet, 0, n, n + m - 1, '.XLSX', fmt)
            set_column_name(worksheet, DOCX_COLUMNS, DOCX_COLUMN_INDEXES, 1, 0, fmt)
            set_column_name(worksheet, XLSX_COLUMNS, XLSX_COLUMN_INDEXES, 1, n, fmt)
            set_column_width(worksheet, 0, DOCX_COLUMNS, DOCX_COLUMN_INDEXES, DOCX_COLUMNS_WIDTH)
            set_column_width(worksheet, n, XLSX_COLUMNS, XLSX_COLUMN_INDEXES, XLSX_COLUMNS_WIDTH)
            row = 2
            doc_to_excel = sorted(matches[match_type].items(),
                                  key=lambda x: docx[x[0]][SORT_FIELD] or MIN_SORT_FILED_VALUE)
            for docx_row, xlsx_rows in doc_to_excel:
                xlsx_rows.sort(key=lambda x: xlsx[x][SORT_FIELD] or MIN_SORT_FILED_VALUE)
                k = len(xlsx_rows)
                if k > 1:
                    for i in range(n):
                        merge_rows(worksheet, i, row, row + k - 1,
                                   docx[docx_row][DOCX_COLUMNS[DOCX_COLUMN_INDEXES[i]]['var']], fmt)
                else:
                    write_row(worksheet, docx, docx_row, row, 0, DOCX_COLUMNS, fmt)
                for xlsx_row in xlsx_rows:
                    write_row(worksheet, xlsx, xlsx_row, row, n,
                              XLSX_COLUMNS, fmt)
                    row += 1
                    if xlsx_row in remained_xlsx_rows:
                        remained_xlsx_rows.remove(xlsx_row)
                if docx_row in remained_docx_rows:
                    remained_docx_rows.remove(docx_row)
        worksheet = workbook.get_worksheet_by_name(NO_MATCH_SHEET_LABEL)
        row = 2
        for docx_row in sorted(remained_docx_rows, key=lambda x: docx[x][SORT_FIELD] or MIN_SORT_FILED_VALUE):
            write_row(worksheet, docx, docx_row, row, 0, DOCX_COLUMNS, fmt)
            row += 1
        for xlsx_row in sorted(remained_xlsx_rows, key=lambda x: xlsx[x][SORT_FIELD] or MIN_SORT_FILED_VALUE):
            write_row(worksheet, xlsx, xlsx_row, row, n, XLSX_COLUMNS, fmt)
            row += 1

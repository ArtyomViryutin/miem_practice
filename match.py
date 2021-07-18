#!/usr/bin/env python3
from file_parser import parse_docx, parse_xls
from tools import find_row_matches, dump_to_excel
import argparse
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description='Парсинг .doc(x) и .xls(x)')
    parser.add_argument('-p', '--path', required=True, metavar=None,
                        help='Абсолютный или относительный путь к папке с .doc(x) и .xls(x) файлами')
    parser.add_argument('-up', '--upload-path', default=Path(__file__).parent.absolute(), required=False,
                        help=r'Путь для сохранения результата: C:\Users\Ivan\Рабочий стол')
    parser.add_argument('-f', '--filename', required=False,
                        help='Имя результирующего файла с совпадениями: filename.xls или filename.xlsx',
                        default='matches.xlsx')
    args = parser.parse_args()
    path = Path(args.path)
    path = Path.resolve(path)
    docx_rows, xlsx_rows = {}, {}
    for child in path.iterdir():
        suffix = child.suffix
        if suffix in ('.doc', '.docx'):
            docx_rows.update(parse_docx.parse(child))
        elif suffix in ('.xlsx', '.xls'):
            xlsx_rows.update(parse_xls.parse(child))
    matches = find_row_matches(docx_rows, xlsx_rows)
    upload_path = Path(args.upload_path) / args.filename
    dump_to_excel(matches, docx_rows, xlsx_rows, upload_path)


if __name__ == '__main__':
    main()



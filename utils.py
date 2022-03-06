from os import path
import re


def get_file_id(filepath):
    head, _ = path.split(filepath)
    return path.split(head)[1]


def id_cols(sheet, col_cord = 'A'):
    cols = sheet[col_cord]
    for cell in cols:
        if cell.value is not None and ALL_NUMBER.match(str(cell.value)):
            yield cell

ALL_NUMBER = re.compile(r'\d+')

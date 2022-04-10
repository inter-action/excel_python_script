from os import path
import re


def get_file_id(filepath):
    head, _ = path.split(filepath)
    head = path.split(head)[1]
    matched = ALL_NUMBER.match(head)
    if matched is None:
        raise Exception('id is not found')
    id = head[matched.start():matched.end()]
    return id


def id_cols(sheet, col_cord = 'A'):
    cols = sheet[col_cord]
    for cell in cols:
        if cell.value is not None and ALL_NUMBER.match(str(cell.value)):
            yield cell

ALL_NUMBER = re.compile(r'\d+')

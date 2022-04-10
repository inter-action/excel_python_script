# coding=utf-8
# encoding
# https://www.python.org/dev/peps/pep-0263/

# api reference: 
#    https://openpyxl.readthedocs.io/en/stable/_modules/index.html
#    https://openpyxl.readthedocs.io/en/3.1/_modules/openpyxl/workbook/workbook.html#Workbook

# cheatsheet: 
#   https://www.pythoncheatsheet.org/#Lists
#   https://perso.limsi.fr/pointal/_media/python:cours:mementopython3-english.pdf

import traceback
from pprint import pprint
import model
import excel
import config


def parse_data(sheet):
    target_idx = []
    for row in sheet.iter_rows(min_row=0, max_row=100, min_col=0, max_col=7):
        for col in row:
            if (col.value == '健管中心'):
                target_idx.append(col.row)
            if (col.value == '科技公司'):
                target_idx.append(col.row)
        # why ?
        if len(target_idx) == 2 and row[0] == '': break

    return [parse_section(sheet, idx) for idx in target_idx]


def parse_section(sheet, row_idx):
    row = row_idx + 1
    id = sheet.cell(row, 1).value
    total = sheet.cell(row, 2).value

    loc = get_loc(row_idx, 1)
    rows = []
    for row in sheet.iter_rows(min_row=loc.min_row, max_row=loc.max_row, min_col=loc.min_column, max_col=loc.max_column):
        values = []
        for col in row:
            values.append(get_cell_value(col))
        rows.append(values)


    idx_group = get_group_row(sheet)

    if idx_group != -1:
        for row in sheet.iter_rows(min_row=idx_group, max_row=idx_group, min_col=3, max_col=7):
            values = []
            for col in row:
               values.append(get_cell_value(col)) 
            rows.append(values)

    return model.DataClassCard(id, total, rows)


def get_group_row(sheet):
    for row in sheet.iter_rows(min_row=0, max_row=100, min_col=3, max_col=3):
        for col in row:
            if not isinstance(col.value, str):
                continue
            if col.value.strip() == '合计':
                return -1
            elif col.value.strip() == '职团渠道':
                return col.row


def get_cell_value(col):
    if col.data_type != 'f':
        return col.value
    else:
        return None

def write_card_to_sheet(sheet, card, index):
    # move one down
    total_cell = sheet.cell(index, 2)
    total_cell.value = card.total

    loc = get_loc(index, 0)
    for row_idx, row in enumerate(
            sheet.iter_rows(loc.min_row, loc.max_row, loc.min_column, loc.max_column)):
        if row_idx < len(card.rows):
            for col_idx, cell in enumerate(row):
                t_value = card.rows[row_idx][col_idx]
                if col_idx == 0:
                    if cell.value is None or t_value is None:
                        break;
                    elif str(t_value).strip() == str(cell.value).strip():
                        continue 
                    else:
                        break;
                
                # if t_value is not None:
                cell.value = t_value
                


def open_to_write(sheet, cards):
    id_to_row_index_map = {}
    card_map = {}
    for card in cards:
        id = card.id
        id_to_row_index_map[id] = None
        card_map[id] = card

    # first row in summary sheet
    for cell in sheet['A']:
        if cell.value in id_to_row_index_map and id_to_row_index_map[cell.value] == None:
            id_to_row_index_map[cell.value] = cell.row

    print('result index is', id_to_row_index_map)

    for key, value in id_to_row_index_map.items():
        write_card_to_sheet(sheet, card_map[key], id_to_row_index_map[key])


def get_loc(row_idx, row_offset):
    [min_row, max_row] = [row_idx + row_offset, row_idx + 3 + row_offset]
    return model.Loc(min_row, max_row, 3, 7)




def all_cards_in_dir(dir_path, sub_sheet_name):
      for entry in excel.excel_files(dir_path):
        wb = None
        try:
            [pathname, filename] = entry
            print(f'process file: {pathname}')
            wb = excel.get_workbook(pathname, True)
            ws = excel.get_working_sheet(wb, sub_sheet_name)
            if ws is None:
                print(f'\n\n[WARN] no sheet found for name: {sub_sheet_name}, file: {pathname}\n sheetnames: {wb.sheetnames}')
                continue
            working_title = ws.title
            
            print(f'filename: {filename} sheet name: {working_title}')

            result = parse_data(ws)

            print(f'sheet result is:\n ${result}')

            if (len(result) != 2):
                print(f'no matching data found in {pathname}-> {working_title}')

            yield result
        finally:
            if wb is not None: wb.close()


task_entries = [
    ('2022年应收账款账龄分析表', '2022年应收账款账龄分析表'),
    ('2021年应收账款账龄分析表', '2021年同期应收账款账龄分析表')
]
target_wb = excel.get_workbook(config.SUMMARY_SHEET_FILE, False)
try:
    for sum_sheet_name, sub_sheet_name in task_entries: 
        target_ws = excel.get_working_sheet(target_wb, sum_sheet_name)
        for cards in all_cards_in_dir(config.FOLDER, sub_sheet_name):
            if len(cards) == 2:
                print(f'parse_data 1: {str(cards[0])}')
                print(f'parse_data 2: {str(cards[1])}')
                open_to_write(target_ws, cards)

    target_wb.save(filename=config.dist("sumary.xlsx"))
except Exception as inst:
    traceback.print_exc()
    print('got undefined error: ', inst, inst.__cause__, inst.__traceback__)
finally:
    if target_wb is not None: target_wb.close()

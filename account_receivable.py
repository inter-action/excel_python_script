from statistics import mode
import traceback
import re
from enum import Enum
import typing

import excel
import utils
import model
import config
import os

class AccountReceivable:
    def __init__(self, id, cols, filename):
        self.rows = cols
        self.id = id
        self.filename = filename


    def __str__(self):
        return "AccountReceivable({}, {})".format(self.id, self.rows)

    def __repr__(self) -> str:
        return "AccountReceivable({}, {})".format(self.id, self.rows)

class AccountReceivableByYear:
    def __init__(self):
        self.items = {}
    
    def add_rows(self, item_type: model.REVENUE_TYPES, datas: typing.List[AccountReceivable]):
        if self.items.get(item_type.value) is None:
            self.items[item_type.value] = []
        self.items[item_type.value].extend(datas)

    # suppose to be 2021
    def a_rows(self):
        return self.items.get(model.REVENUE_TYPES.A.value)

    # suppose to be 2021
    def b_rows(self):
        return self.items.get(model.REVENUE_TYPES.B.value)


    def __repr__(self) -> str:
        return "AccountReceivableByYear({})".format(self.items)


class IdRange:
    def __init__(self):
        self.range = {}

    def set_range(self, id, row):
        row_end = row + 3
        key = str(id)
        if not key in self.range:
            self.range[key] = "C{}:C{}".format(row, row_end)
            return True
        return False

    def get_range(self, id):
        return self.range.get(id)

    def values(self):
        return self.range.values()

    def items(self):
        return self.range.items()

    def __str__(self):
        return "{}".format(self.range)


#---------- parse logic

def parse(wb, account_by_year: AccountReceivableByYear, filename: str):
    names = source_sheet_names(wb)
    for name in names:
        print(f'parsed sub sheet name: {name}, filename: {filename}')
        ws = excel.get_working_sheet(wb, name)
        if ws is None: continue

        account_recivable = parse_by_sheet(ws, filename)
        if model.REVENUE_TYPES.A.value in name:
            account_by_year.add_rows(model.REVENUE_TYPES.A, account_recivable)
        if model.REVENUE_TYPES.B.value in name:
            account_by_year.add_rows(model.REVENUE_TYPES.B, account_recivable)
    

def parse_by_sheet(sheet, filename):
    ranges = parse_source_id_range(sheet)
    for id, range in ranges.items():
        cols_result = []
        for cols in excel.iter_cols_by_range(sheet, range):
            one_col = []
            for cell in cols:
                one_col.append(cell.value)
            cols_result.append(one_col)

        yield AccountReceivable(id, cols_result[0], filename)

def parse_source_id_range(target_sheet):
    range = IdRange()
    cols = list(utils.id_cols(target_sheet))[0:2]
    for cell in cols:
        range.set_range(cell.value, cell.row)
    return range


SOURCE_SHEET_NAME_REG = re.compile(r'\d{4}年\d+月以内应收账款')
def source_sheet_names(wb):
    result = list(filter(lambda x: SOURCE_SHEET_NAME_REG.match(x), wb.sheetnames)) 
    return result



def all_account_receivable(folder):
    account_receivable_by_year = AccountReceivableByYear()
    for filepath, filename in excel.excel_files(folder):
        wb = excel.get_workbook(filepath, True)
        try:
            parse(wb, account_receivable_by_year, filepath)
        finally:
            wb.close()
    return account_receivable_by_year




#---------- write logic


SUM_SHEET_NAME_REG = re.compile(r'\d{4}年\d+月以内应收账款')
def sum_sheet_names(wb):
    result = list(filter(lambda x: SUM_SHEET_NAME_REG.match(x), wb.sheetnames)) 
    return result


def parse_write_range(sum_sheet):
    range = IdRange()
    for cell in utils.id_cols(sum_sheet):
        range.set_range(cell.value, cell.row)
    return range

def write_to_sheet(sum_sheet, write_range: IdRange, datas: typing.List[AccountReceivable]):
    for account_receivable in datas:
        range = write_range.get_range(account_receivable.id)
        if range is None:
            print(f'range not found for in sheet: {sum_sheet}, id: {account_receivable.id}, file: {account_receivable.filename}')
        else:
            for col in excel.iter_cols_by_range(sum_sheet, range):
                for idx, cell in enumerate(col):
                    cell.value = account_receivable.rows[idx]


def write_all(target_file, save_filename, account_by_year: AccountReceivableByYear):
    wb_target = excel.get_workbook(target_file, True)
    try:
        for sheetname in sum_sheet_names(wb_target):
            print(f'sum sheet name: {sheetname}')
            ws_target = excel.get_working_sheet(wb_target, sheetname)
            range = parse_write_range(ws_target)
            if model.REVENUE_TYPES.A.value in sheetname:
                write_to_sheet(ws_target, range, account_by_year.a_rows())
            if model.REVENUE_TYPES.B.value in sheetname:
                write_to_sheet(ws_target, range, account_by_year.b_rows())
        wb_target.save(filename=save_filename)
    except Exception as inst:
        traceback.print_exc()
        print('got undefined error: ', inst, inst.__cause__, inst.__traceback__)
    finally:
        wb_target.close()
        


if __name__ == '__main__':
    # wb = excel.get_workbook('1月/7101/7101应收账款账龄分析表-22.1月.xlsx', True)
    # ws = excel.get_working_sheet(wb, '2022年1月应收账款')
    # parse_result = list(parse(ws))
    # src_names = source_sheet_names(wb)
    # print(f'{src_names}')


    # wb_sum = excel.get_workbook('1月/应收账款账龄分析表-1月-基础表.xlsx', True)
    # print(f'{sum_sheet_names(wb_sum)}')

    # ws_sum = excel.get_working_sheet(wb_sum, '2022年12月营收')

    SAVE_RESULT_FILENAME = config.dist('account_receivable.xlsx')
    account_receivable_by_year = all_account_receivable(config.FOLDER)
    write_all(config.SUMMARY_SHEET_FILE, SAVE_RESULT_FILENAME, account_receivable_by_year)
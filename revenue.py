import traceback
import re
from enum import Enum

import excel
import utils

RANGE = 'A2:C8'

# this related to `first_col` param of write_unit function
class REVENUE_TYPES(Enum):
    A = '2022'
    B = '2021'

    def is_a(self):
        return self == REVENUE_TYPES.A

class Revenue:
    def __init__(self, id, cols):
        self.rows = cols
        self.id = id

    # 2022
    def col_b(self):
        return self.rows[1]

    # 2021
    def col_c(self):
        return self.rows[2]


class WriteRange:
    def __init__(self):
        self.range = {}

    def set_range(self, id, row):
        row_end = row + 6
        self.range[str(id)] = "C{}:C{}".format(row, row_end)

    def get_range(self, id):
        return self.range[id]

    def __str__(self):
        return "{}".format(self.range)

def parse(sheet, id, range = RANGE):
    loc = excel.range_to_loc(range)
    cols_result = []
    for cols in sheet.iter_cols(loc.min_column, loc.max_column, loc.min_row, loc.max_row):
        one_col = []
        for cell in cols:
            one_col.append(cell.value)
        cols_result.append(one_col)

    return Revenue(id, cols_result)


def write_unit(target_sheet, target_range: WriteRange, revenue: Revenue, first_col = True):
    """
    :first_col: usually decide we take 2022 col or 2021 cols as datasource.
    """
    range = target_range.get_range(revenue.id)
    if range is not None:
        for cols in excel.iter_cols_by_range(target_sheet, range):
            i = 0
            for cell in cols:
                if first_col:
                    cell.value = revenue.col_b()[i]
                else:
                    cell.value = revenue.col_c()[i]
                print(f'write cell: {cell}, with value {cell.value}')
                i += 1
    else:
        print(f'[WARN] range not found for id: {revenue.id}')




ALL_NUMBER = re.compile(r'\d+')

def parse_write_range(target_sheet):
    range = WriteRange()
    cols = target_sheet['A']
    for cell in cols:
        if cell.value is not None and ALL_NUMBER.match(str(cell.value)):
            range.set_range(cell.value, cell.row)

    return range


def write_all(target_file, save_filename, all_revenues):
    wb_target = excel.get_workbook(target_file, True)
    try:
        for revenue_type, sheetname in target_revenue_sheetnames(wb_target):
            ws_target = excel.get_working_sheet(wb_target, sheetname)
            print(f'try to write {ws_target}, {sheetname}')
            range = parse_write_range(ws_target)
            for revenue in all_revenues:
                if revenue_type.is_a():
                    write_unit(ws_target, range, revenue, True)
                else:
                    write_unit(ws_target, range, revenue, False)
        wb_target.save(filename=save_filename)
    except Exception as inst:
        traceback.print_exc()
        print('got undefined error: ', inst, inst.__cause__, inst.__traceback__)
    finally:
        wb_target.close()
        


def target_revenue_sheetnames(workbook):
    keyword = '营收'
    result = list(filter(lambda x: keyword in x, workbook.sheetnames)) 
    print(f'results: {result}')
    if result is not None:
        for name in result:
            if REVENUE_TYPES.A.value in name:
                print(f'results:a {name}')
                yield (REVENUE_TYPES.A, name)
            if REVENUE_TYPES.B.value in name:
                print(f'results:b {name}')
                yield (REVENUE_TYPES.B, name)


def all_revenues(folder, sheetname):
    for filepath, _ in excel.excel_files(folder):
        wb = excel.get_workbook(filepath, True)
        try:
            ws = excel.get_working_sheet(wb, sheetname)
            file_id = utils.get_file_id(filepath)
            revenue = parse(ws, file_id)
            yield revenue
        finally:
            wb.close()




if __name__ == '__main__':
    # wb = excel.get_workbook('1月/7101/7101应收账款账龄分析表-22.1月.xlsx', True)
    # ws = excel.get_working_sheet(wb, '分渠道营收统计')
    # revenue = parse(ws, 7101)
    # print(f'2022: {revenue.col_b()}')
    # print(f'2021: {revenue.col_c()}')

    # wb_sum = excel.get_workbook('1月/应收账款账龄分析表-1月-基础表.xlsx', True)
    # ws_sum = excel.get_working_sheet(wb_sum, '2022年12月营收')
    # range = parse_write_range(ws_sum)

    # write_unit(ws_sum, range, revenue)
    # print(f'range: {range}')

    # wb_sum.save(filename='revenue.xlsx')

    # names = revenue_sheetnames(wb_sum)
    # print(f'names: {names}')


    # wb_sum = excel.get_workbook(TARGET_SHEET_NAME, True)
    # ws_sum = excel.get_working_sheet(wb_sum, '2022年12月营收')

    # range = parse_write_range()
    # names = revenue_sheetnames(wb_sum)

    TARGET_FILE_NAME = '1月/应收账款账龄分析表-1月-基础表.xlsx'
    FOLDER = '1月'
    save_result_filename = 'revenue.xlsx'
    revenues = list(all_revenues(FOLDER, '分渠道营收统计'))
    write_all(TARGET_FILE_NAME, save_result_filename, revenues)




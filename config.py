

import os

SUMMARY_SHEET_FILE='./3月/应收账款账龄分析表-2月-基础表.xlsx'
FOLDER='9月'
DIST='dist'


def dist(filename):
    return os.path.join(DIST, filename)
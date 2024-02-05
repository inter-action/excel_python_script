

import os

SUMMARY_SHEET_FILE='./应收账款账龄分析表-2024年1月.xlsx'
FOLDER='1月'
DIST='dist'


def dist(filename):
    return os.path.join(DIST, filename)
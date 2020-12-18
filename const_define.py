# -*- coding: UTF-8 -*-
import os

STATISTICAL_TABLE_KEYWORD = '統籌統計表'
DATA_FOLDER_NAME = 'data'
SALARY_DATA_FOLDER_NAME = os.path.join(os.path.join(os.getcwd(), DATA_FOLDER_NAME), 'salary')


DETAILED_LEDGER_FILENAME = r'107年方殷營業支出.xlsx'
DATA_FOLDER_FULL_PATH = os.path.join(os.getcwd(), DATA_FOLDER_NAME)
DETAILED_LEDGER_FULL_PATH = os.path.join(DATA_FOLDER_FULL_PATH, DETAILED_LEDGER_FILENAME)
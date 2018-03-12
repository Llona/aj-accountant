# -*- coding: UTF-8 -*-
import openpyxl
import shutil
import os
import const_define


# active sheet name
# print(workbook.active)

# load excel file
workbook = openpyxl.load_workbook(const_define.DETAILED_LEDGER_FULL_PATH)
# get all sheet name
worksheets = workbook.get_sheet_names()

# get sheet content
sheet = workbook.get_sheet_by_name(worksheets[0])
# print(sheet)
# print(sheet.title)
# print(sheet.cell(row=2, column=2).value)

for rowOfCell in sheet['B2':'G2']:
    print(rowOfCell)
    for cell in rowOfCell:
        # print(cell.coordinate, cell.value)
        print(cell.value)

# shutil.copyfile(os.path.join('T:'), os.path.join('ttt.txt'))
# filename_netdriver = os.path.join(r"t:", 'vv')
# filename_netdriver = os.path.join(filename_netdriver, 'Roy')
# filename_netdriver = os.path.join(filename_netdriver, 'command.txt')
# print(filename_netdriver)
# shutil.copy(filename_netdriver, os.getcwd())

filename_netdriver = os.path.join(r"t:", r'vv\Roy\command.txt')
print(filename_netdriver)
shutil.copy(filename_netdriver, os.getcwd())

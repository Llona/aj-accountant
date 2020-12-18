# -*- coding: UTF-8 -*-
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
import os
import const_define
import re
from decimal import Decimal, ROUND_HALF_UP


class PerformanceCalculation(object):
    def __init__(self):
        self.salary_folder_filename = const_define.SALARY_DATA_FOLDER_NAME
        self.sheet = None
        self.sheet_formula = None
        self.sheet_statistical_table = None
        self.name_mapping_dic = {}
        self.overall_bonus_dic = {}
        self.total_bonus_dic = {}

        self.get_name_mapping_dic()

    def get_name_mapping_dic(self):
        with open('name.db', 'r', encoding='UTF-8') as file_h:
            lines = file_h.readlines()
            for line in lines:
                line_name = line.split(':')
                chinese_name = line_name[0].rstrip('\n\r')
                english_name = line_name[1].rstrip('\n\r')
                self.name_mapping_dic[english_name] = chinese_name

    def calc_all_salary(self):
        performance_files = self.get_all_perf_file_name()
        self.get_statistical_table_dic()
        # print(self.overall_bonus_dic)
        for full_filepath in performance_files:
            filename_full = os.path.basename(full_filepath)
            filename_with_ext = os.path.splitext(filename_full)
            filename = filename_with_ext[0]
            file_ext = filename_with_ext[1]

            if filename.find(const_define.STATISTICAL_TABLE_KEYWORD) < 0 and file_ext == '.xlsx':
                name = self.get_name_from_filename(filename).lower()
                print(name)
                personal_bonus = self.calc_salary(full_filepath)
                print(personal_bonus)
                print(self.overall_bonus_dic[name])
                total_bonus = self.overall_bonus_dic[name] + int(personal_bonus)

                self.total_bonus_dic[name] = total_bonus
        print(self.total_bonus_dic)

    def get_statistical_table_dic(self):
        statistical_table_filename = self.get_statistical_table_filename(self.get_all_perf_file_name())
        print(statistical_table_filename)
        workbook = openpyxl.load_workbook(statistical_table_filename, read_only=True, data_only=True)
        worksheets = tuple(workbook.sheetnames)
        self.sheet_statistical_table = workbook[worksheets[0]]
        for row in self.sheet_statistical_table.iter_rows():
            for cell in row:
                if cell.value == '各別獎金':
                    self.get_overall_value(cell)

    def get_overall_value(self, cell):
        # for name in self.name_mapping_dic.keys():
        for i in range(1, 99):
            name = self.sheet_statistical_table.cell(row=cell.row-1, column=cell.column+i)
            name = name.value
            if name:
                for key, value in self.name_mapping_dic.items():
                    if key == name.lower():
                        self.overall_bonus_dic[key] = self.sheet_statistical_table.cell(row=cell.row+1, column=cell.column+i).value
            else:
                return

    def calc_salary(self, full_filename):
        workbook = openpyxl.load_workbook(full_filename, read_only=True, data_only=True)
        worksheets = tuple(workbook.sheetnames)
        self.sheet = workbook[worksheets[0]]
        workbook_temp = openpyxl.load_workbook(full_filename, read_only=True, data_only=False)
        worksheets_temp = tuple(workbook.sheetnames)
        self.sheet_formula = workbook_temp[worksheets_temp[0]]
        for row in self.sheet_formula.iter_rows():
            for cell in row:
                if cell.value == '個人Total' and cell.column == column_index_from_string('I'):
                    perf_value = self.calculate_value_cell(self.sheet_formula.cell(row=cell.row, column=cell.column+1))
                    return perf_value
        print("-------------------------")

    def calculate_value_cell(self, cell):
        formula = cell.value
        performance_value = 0

        if not self.is_formula(formula):
            return formula

        if formula.find("SUM") >= 0:
            sum_index = formula.split(':')
            sum_min_index = coordinate_from_string(re.sub(r'=SUM\(', '', sum_index[0]))
            sum_max_index = coordinate_from_string(re.sub(r'\)', '', sum_index[1]))
            # print(sum_min_index)
            # print(sum_max_index)
            sum_min_col = column_index_from_string(sum_min_index[0])
            sum_min_row = int(sum_min_index[1])
            sum_max_col = column_index_from_string(sum_max_index[0])
            sum_max_row = int(sum_max_index[1])

            if sum_min_col != sum_max_col:
                print('not sum the same col, not support at this time!')
                return
            row_count = sum_min_row
            # print(sum_min_col, sum_min_row, sum_max_col, sum_max_row)
            # print(row_count)
            while row_count < sum_max_row:
                value = self.sheet.cell(row=row_count, column=sum_min_col).value

                if value:
                    # print(value)
                    # print(type(value))
                    try:
                        float(value)
                        # print(value)
                        value = self.round_v2(value)
                        # print(value)
                        performance_value = performance_value + value
                    except Exception as e:
                        str(e)
                row_count += 1
            return performance_value
        elif formula.find("+") >= 0:
            formula = re.sub('=', '', formula)
            cell_list = formula.split('+')
            for cell_index in cell_list:
                value = self.sheet[cell_index].value

                if value:
                    # print(value)
                    try:
                        float(value)
                        # print(value)
                        value = self.round_v2(value)
                        # print(value)
                        performance_value = performance_value + value
                    except Exception as e:
                        str(e)
            return performance_value
        else:
            print('not SUM formula')

    @staticmethod
    def round_v2(number):
        origin_num = Decimal(number)
        result_num = origin_num.quantize(Decimal('0'), rounding=ROUND_HALF_UP)
        return result_num

    @staticmethod
    def is_formula(string):
        try:
            if string[0] == '=':
                return True
            else:
                return False
        except:
            return False

    @staticmethod
    def get_name_from_filename(filename):
        temp_str = filename.split(" ")[0]
        name = temp_str.split("_")[1]
        return name

    @staticmethod
    def get_statistical_table_filename(filelist):
        for full_filename in filelist:
            file_name_ext = os.path.basename(full_filename)
            filename = os.path.splitext(file_name_ext)[0]
            file_ext = os.path.splitext(file_name_ext)[1]
            if filename.find(const_define.STATISTICAL_TABLE_KEYWORD) > -1 and file_ext == '.xlsx':
                return full_filename
        return None

    def get_all_perf_file_name(self):
        file_list = []
        for file_name in os.listdir(self.salary_folder_filename):
            full_path = os.path.join(self.salary_folder_filename, file_name)
            if os.path.isfile(full_path):
                file_list.append(full_path)
        return file_list


salary_h = PerformanceCalculation()
salary_h.calc_all_salary()
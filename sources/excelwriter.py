import os
from pprint import pprint

import openpyxl


class ExcelWriter:
    """Класс по работе с файлами Excel"""

    def __init__(self, filename, sheet, positions, balance, usd_course):
        self.filename = f'../{filename}.xlsx'
        self.positions = positions
        self.balance = balance
        self.usd_course = usd_course

        # Открытие файла и страницы в файле Excel
        # Если файла не существует, то он создается
        if os.path.exists(f'../{filename}'):
            self.workbook = openpyxl.load_workbook(f'../{filename}')
            self.worksheet = self.workbook[sheet]
        else:
            self.workbook = openpyxl.Workbook()
            self.worksheet = self.workbook.active
            self.worksheet.title = sheet
            self.workbook.save(self.filename)

    def write_names_of_columns(self):
        pass

    def write_table_to_excel(self):
        pass
import os
from pprint import pprint

import openpyxl


class ExcelPortfolio:
    """Класс по работе с файлами Excel"""

    def __init__(self, filename, sheet, portfolio_table):
        self.filename = f'../{filename}'
        self.portfolio_table = portfolio_table

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

        self.init_name_columns()

    def write_balance(self):
        """Запись текущего баланса в портфеле"""
        balance = self.portfolio_table['balance']

        for row, position in zip(self.worksheet.iter_rows(min_row=12, max_row=len(balance)+12, min_col=4, max_col=4), balance.keys()):
            for cell in row:
                cell.value = f'Баланс ({position})'

        for row, position in zip(self.worksheet.iter_rows(min_row=12, max_row=len(balance)+12, min_col=5, max_col=5), balance.values()):
            for cell in row:
                cell.value = position

        self.workbook.save(self.filename)

    def write_positions(self):
        """Запись текущих позиций бумаг портфеля"""
        positions = self.portfolio_table['positions']

        for row, position in zip(self.worksheet.iter_rows(min_row=2, max_row=len(positions)+2, max_col=len(positions[0])), positions):
            for cell, el in zip(row, position):
                cell.value = el

        self.workbook.save(self.filename)

    def init_name_columns(self):
        """Инициализация имен колонок таблицы"""
        names_of_table = self.portfolio_table['names_of_table']

        for col, name in zip(self.worksheet.iter_cols(max_row=1, max_col=len(names_of_table)), names_of_table):
            for cell in col:
                cell.value = name

        self.workbook.save(self.filename)

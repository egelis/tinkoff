import os
from pprint import pprint

import openpyxl


class ExcelPortfolio:
    """Класс по работе с файлами Excel"""

    def __init__(self, filename, sheet, portfolio_table, portfolio_price):
        self.filename = f'../{filename}'
        self.portfolio_table = portfolio_table
        self.portfolio_price = portfolio_price

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

    def write_balance(self, min_row, min_col):
        """
        Запись текущего баланса в портфеле
        Принимает левый верхний угол, с которого начинается заполнение
        """

        balance = self.portfolio_table['balance']

        for row, position in zip(self.worksheet.iter_rows(min_row=min_row, max_row=len(balance) + min_row,
                                                          min_col=min_col, max_col=min_col), balance.keys()):
            for cell in row:
                cell.value = f'Баланс ({position})'

        for row, position in zip(self.worksheet.iter_rows(min_row=min_row, max_row=len(balance) + min_row,
                                                          min_col=min_col + 1, max_col=min_col + 1), balance.values()):
            for cell in row:
                cell.value = position

        self.workbook.save(self.filename)

    def write_positions(self, min_row, min_col):
        """
        Запись текущих позиций бумаг портфеля
        Принимает левый верхний угол, с которого начинается заполнение
        """

        positions = self.portfolio_table['positions']

        for row, position in zip(
                self.worksheet.iter_rows(min_row=min_row, min_col=min_col, max_row=len(positions) + min_row,
                                         max_col=len(positions[0]) + min_col), positions):
            for cell, el in zip(row, position):
                cell.value = el

        self.workbook.save(self.filename)

    def write_names_of_columns(self, min_row, min_col):
        """
        Написание имен колонок таблицы
        Принимает левый верхний угол, с которого начинается заполнение
        """

        names_of_table = self.portfolio_table['names_of_table']

        for col, name in zip(self.worksheet.iter_cols(min_row=min_row, min_col=min_col, max_row=min_row,
                                                      max_col=len(names_of_table) + min_col), names_of_table):
            for cell in col:
                cell.value = name

        self.workbook.save(self.filename)

    def write_portfolio_price_rub(self):
        """
        Написание итоговой цены портфеля
        Принимает левый верхний угол, с которого начинается заполнение
        """

        self.worksheet['A1'] = 'Общая цена портфеля:'
        self.worksheet['B1'] = self.portfolio_price

        self.workbook.save(self.filename)

    def write_share_position(self, min_row, min_col):
        for i, row in enumerate(self.worksheet.iter_rows(min_row=min_row, min_col=min_col,
                                                         max_row=len(self.portfolio_table['positions']) + min_row - 1,
                                                         max_col=min_col), start=min_row):
            for cell in row:
                cell.value = f"=E{i} / B1"

        self.workbook.save(self.filename)

    def write_table_to_excel(self):
        self.write_names_of_columns(3, 1)
        self.write_positions(4, 1)
        self.write_balance(len(self.portfolio_table['positions']) + 4 + 1, 4)
        self.write_share_position(4, 6)
        self.write_portfolio_price_rub()

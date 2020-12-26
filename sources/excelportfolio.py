import os
from pprint import pprint

import openpyxl


class ExcelPortfolio:
    """Класс по работе с файлами Excel"""

    def __init__(self, filename, sheet):
        self.filename = f'../{filename}'

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

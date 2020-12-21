from pprint import pprint
from decimal import Decimal
import os

from tinkoffapi import TinkoffApi
from excelportfolio import ExcelPortfolio
from creator import TableCreator

if __name__ == "__main__":
    tinkoff = TinkoffApi()
    positions = tinkoff.get_portfolio_positions()
    balance = tinkoff.get_portfolio_balance()
    usd_course = tinkoff.get_usd_course()

    converter = TableCreator(positions, balance, usd_course)
    portfolio = converter.get_portfolio_table_for_excel()
    pprint(portfolio)

    excel_file = ExcelPortfolio('Инвест профиль.xlsx', 'Лист1', portfolio)
    excel_file.write_positions()
    excel_file.write_balance()

    os.system(f'start excel.exe \"../Инвест профиль.xlsx\""')

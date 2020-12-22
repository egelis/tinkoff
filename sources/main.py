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
    portfolio_price = converter.get_portfolio_price_rub()

    excel_file = ExcelPortfolio('Инвест профиль.xlsx', 'Лист1', portfolio, portfolio_price)
    excel_file.write_names_of_columns(3, 1)
    excel_file.write_positions(4, 1)
    excel_file.write_balance(len(portfolio['positions']) + 4 + 1, 4)
    excel_file.write_portfolio_price_rub()

    os.system(f'start excel.exe \"../Инвест профиль.xlsx\""')

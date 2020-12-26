from pprint import pprint
from decimal import Decimal
import os

from tinkoffapi import TinkoffApi
from excelwriter import ExcelWriter

if __name__ == "__main__":
    tinkoff = TinkoffApi()
    positions = tinkoff.get_portfolio_positions()
    balance = tinkoff.get_portfolio_balance()
    usd_course = tinkoff.get_usd_course()
    eur_course = tinkoff.get_eur_course()

    pprint(positions)
    pprint(balance)
    pprint(usd_course)
    pprint(eur_course)

    # excel_writer = ExcelWriter('Инвест профиль', 'Лист1', positions, balance, usd_course)
    # excel_writer.write_table_to_excel()

    # os.system(f'start excel.exe \"../Инвест профиль.xlsx\""')

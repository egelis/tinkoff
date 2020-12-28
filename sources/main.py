import os

from tinkoffapi import TinkoffApi
from excelwriter import ExcelWriter

if __name__ == "__main__":
    tinkoff = TinkoffApi()
    positions = tinkoff.get_portfolio_positions()
    balance = tinkoff.get_portfolio_balance()
    courses = {'USD': tinkoff.get_usd_course(), 'EUR': tinkoff.get_eur_course()}

    excel_writer = ExcelWriter('Инвест профиль', 'Инвестиции', positions, balance, courses)
    excel_writer.write_table_to_excel()

    print('COMPLETE!')
    os.system(f'start excel.exe \"../Инвест профиль.xlsx\""')

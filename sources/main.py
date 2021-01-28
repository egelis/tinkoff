import os

from tinkoffapi import TinkoffApi
from excelwriter import ExcelWriter

if __name__ == "__main__":
    tinkoff = TinkoffApi()
    operations = [tinkoff.get_portfolio_operations(), tinkoff.get_candle_from_date]
    positions = tinkoff.get_portfolio_positions()
    balance = tinkoff.get_portfolio_balance()
    courses = {'USD': tinkoff.get_usd_course(), 'EUR': tinkoff.get_eur_course()}

    excel_writer = ExcelWriter('Инвест профиль', 'Инвестиции', positions, balance, courses, operations)
    excel_writer.write_table_to_excel()

    print('COMPLETE!')
    # os.system(f'start excel.exe \"../Инвест профиль.xlsx\""')

# except tinvest.exceptions.UnexpectedError:

from pprint import pprint
from decimal import Decimal
import os

from tinkoffapi import TinkoffApi

if __name__ == "__main__":
    tinkoff = TinkoffApi()
    positions = tinkoff.get_portfolio_positions()
    balance = tinkoff.get_portfolio_balance()
    usd_course = tinkoff.get_usd_course()

    os.system(f'start excel.exe \"../Инвест профиль.xlsx\""')

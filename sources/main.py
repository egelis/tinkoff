from pprint import pprint
from decimal import Decimal
import os

from tinkoffapi import TinkoffApi
from excel import Excel


def get_currency_symbol(currency) -> str:
    if currency == 'RUB':
        return '₽'
    elif currency == 'USD':
        return '$'


def get_price_position_rub(position, usd_course) -> float:
    """Возврат общей цены позиции на данный момент в рублях"""
    currency = position.average_position_price.currency.value
    price = 0
    if currency == 'RUB':
        price = position.expected_yield.value + position.average_position_price.value \
                * position.balance
    elif currency == 'USD':
        price = (position.expected_yield.value + position.average_position_price.value
                 * position.balance) * usd_course

    return round(price, 2)


def get_portfolio_for_excel(positions, balance, usd_course) -> dict:
    """Преобразование позиций в удобный для обработки класса Excel словарь"""
    portfolio = {'names_of_table': ('Название', 'Котировка', 'Цена, шт.', 'Кол-во', 'Общая цена, руб.', '%'),
                 'positions': []}

    for position in positions:
        exceptions = ('Доллар США', 'Евро')
        if position.name not in exceptions:
            total_price = get_price_position_rub(position, usd_course)
            unit_price = round(position.average_position_price.value + position.expected_yield.value / position.balance,
                               2)
            portfolio['positions'].append((position.name,
                                           position.ticker,
                                           f'{unit_price}{get_currency_symbol(position.average_position_price.currency.value)}',
                                           int(position.balance),
                                           f'{total_price}₽'))

    portfolio['balance'] = {}
    for position in balance:
        portfolio['balance'][position.currency.value] = position.balance

    return portfolio


def get_portfolio_price(positions, balance, usd_course) -> float:
    price = 0
    for position in positions:
        price += get_price_position_rub(position, usd_course)
    price += balance[0].balance

    return round(price, 2)


if __name__ == "__main__":
    tinkoff = TinkoffApi()
    positions, balance = tinkoff.get_portfolio_positions_and_balance()
    usd_course = tinkoff.get_usd_course()
    portfolio = get_portfolio_for_excel(positions, balance, usd_course)

    excel_file = Excel('Инвест профиль.xlsx', 'Лист1', portfolio)
    excel_file.write_positions()
    excel_file.write_balance()

    os.system(f'start excel.exe \"../Инвест профиль.xlsx\""')

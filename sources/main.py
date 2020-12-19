from pprint import pprint
from decimal import Decimal

from tinkoffapi import TinkoffApi


def get_portfolio_positions_for_excel(positions, balance, usd_course):
    """Преобразование позиций в удобный для Excel список"""
    positions_rows = [('Название', 'Котировка', 'Цена, шт.', 'Кол-во', 'Общая цена, руб.', '%')]

    for position in positions:
        exceptions = ('Доллар США', 'Евро')
        if position.name not in exceptions:
            currency = position.average_position_price.currency.value
            total_price = 0

            if currency == 'RUB':
                total_price = position.expected_yield.value + position.average_position_price.value \
                              * position.balance
            elif currency == 'USD':
                total_price = round((position.expected_yield.value + position.average_position_price.value
                                     * position.balance) * usd_course, 2)

            positions_rows.append((position.name,
                                   position.ticker,
                                   round(position.average_position_price.value + position.expected_yield.value /
                                         position.balance, 2),
                                   int(position.balance),
                                   total_price))

    positions_rows.append(('Баланс (RUB)', balance[0].balance))
    positions_rows.append(('Баланс (USD)', balance[1].balance))

    return positions_rows


if __name__ == "__main__":
    tinkoff = TinkoffApi()
    positions, balance = tinkoff.get_portfolio_positions_and_balance()
    portfolio = get_portfolio_positions_for_excel(positions, balance, tinkoff.get_usd_course())

    pprint(portfolio)
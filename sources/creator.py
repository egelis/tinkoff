from pprint import pprint


class TableCreator:
    """Создатель удобной таблицы для добавления в эксель"""

    def __init__(self, positions, balance, usd_course):
        self.positions = positions
        self.balance = balance
        self.usd_course = usd_course

    def get_portfolio_price(self) -> float:
        price = 0
        for position in self.positions:
            price += self.get_price_position_rub(position)
        price += self.balance[0].balance

        return round(price, 2)

    @staticmethod
    def get_currency_symbol(currency) -> str:
        if currency == 'RUB':
            return '₽'
        elif currency == 'USD':
            return '$'
        elif currency == 'EUR':
            return '€'

    def get_price_position_rub(self, position) -> float:
        """Возврат общей цены позиции на данный момент в рублях"""
        currency = position.average_position_price.currency.value
        price = 0
        if currency == 'RUB':
            price = position.expected_yield.value + position.average_position_price.value \
                    * position.balance
        elif currency == 'USD':
            price = (position.expected_yield.value + position.average_position_price.value
                     * position.balance) * self.usd_course

        return round(price, 2)

    @staticmethod
    def get_names_of_table() -> tuple:
        return 'Название', 'Котировка', 'Цена, шт.', 'Кол-во', 'Общая цена, руб.', '%'

    def get_positions_for_table(self) -> list:
        result = []
        for position in self.positions:
            pprint(position)
            exceptions = ('Доллар США', 'Евро')
            if position.name not in exceptions:
                total_price = self.get_price_position_rub(position)
                unit_price = round(
                    position.average_position_price.value + position.expected_yield.value / position.balance,
                    2)
                result.append((position.name,
                               position.ticker,
                               f'{unit_price} {self.get_currency_symbol(position.average_position_price.currency.value)}',
                               int(position.balance),
                               f'{total_price}₽'))

        return result

    def get_balance_for_table(self) -> dict:
        result = {}
        for position in self.balance:
            result[position.currency.value] = f'{position.balance} {self.get_currency_symbol(position.currency.value)}'

        return result

    def get_portfolio_table_for_excel(self) -> dict:
        """
        Преобразование позиций в удобный для обработки класса Excel словарь, включающий имена столбцов,
        позиций бумаг, баланс в ₽, $, €
        """

        portfolio = {}
        portfolio['names_of_table'] = self.get_names_of_table()
        portfolio['positions'] = self.get_positions_for_table()
        portfolio['balance'] = self.get_balance_for_table()

        return portfolio

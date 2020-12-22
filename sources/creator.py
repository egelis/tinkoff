from pprint import pprint


class TableCreator:
    """Создатель удобной таблицы для добавления в эксель"""

    def __init__(self, positions, balance, usd_course):
        self.positions = positions
        self.balance = balance
        self.usd_course = usd_course

    def get_portfolio_price_rub(self) -> str:
        price = 0
        for position in self.positions:
            price += self.get_total_position_price_rub(position)
        price += self.balance[0].balance  # Рублевый баланс

        return f"{round(price, 2)} {self.get_currency_symbol('RUB')}"

    @staticmethod
    def get_currency_symbol(currency) -> str:
        if currency == 'RUB':
            return '₽'
        elif currency == 'USD':
            return '$'
        elif currency == 'EUR':
            return '€'

    def get_total_position_price_rub(self, position) -> float:
        current_currency = position.average_position_price.currency
        price = 0
        if current_currency == 'RUB':
            price = position.average_position_price.value * position.balance + position.expected_yield.value
        elif current_currency == 'USD':
            price = (position.average_position_price.value
                     * position.balance + position.expected_yield.value) * self.usd_course

        return round(price, 2)

    @staticmethod
    def get_names_of_table() -> tuple:
        return 'Название', 'Котировка', 'Цена, шт.', 'Кол-во', 'Общая цена, руб.', '%'

    def get_positions_for_table(self) -> list:
        result = []

        for position in self.positions:
            if position.name in ('Доллар США', 'Евро'):
                continue
            total_position_price_rub = f"{self.get_total_position_price_rub(position)} {self.get_currency_symbol('RUB')}"
            unit_price = round(position.average_position_price.value + (position.expected_yield.value /
                                                                        position.balance), 2)
            # Имя, тикер, цена за шт., кол-во, итоговая цена в руб., процент. составляющая в портфеле
            result.append((position.name,
                           position.ticker,
                           f'{unit_price} {self.get_currency_symbol(position.average_position_price.currency)}',
                           position.balance,
                           total_position_price_rub))

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
        portfolio['portfolio_price'] = self.get_portfolio_price_rub()

        return portfolio

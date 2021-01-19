import os
from decimal import Decimal

from dotenv import load_dotenv
import tinvest


class TinkoffApi:
    """Обёртка для работы с API Тинькова на основе библиотеки tinvest"""

    def __init__(self):
        self.__init_env_var()

        self.__client = tinvest.SyncClient(os.getenv('TINKOFF_API_TOKEN'))
        self.__account_id = os.getenv("TINKOFF_BROKER_ACCOUNT")

    @staticmethod
    def __init_env_var():
        """Инициализация переменных окружения"""
        dotenv_path = os.path.join(os.path.dirname(__file__), '..\.env')
        if os.path.exists(dotenv_path):
            load_dotenv(dotenv_path=dotenv_path)

    def get_usd_course(self) -> Decimal:
        """Получение курса USD"""
        return self.__client.get_market_orderbook(figi="BBG0013HGFT4", depth=1).payload.last_price

    def get_eur_course(self) -> Decimal:
        """Получение курса EUR"""
        return self.__client.get_market_orderbook(figi="BBG0013HJJ31", depth=1).payload.last_price

    def get_portfolio_positions(self) -> list:
        """Получение всех позиций портфеля"""
        return self.__client.get_portfolio(broker_account_id=self.__account_id) \
            .payload.positions

    def get_portfolio_balance(self) -> list:
        """Получение баланса портфеля во всех валютах"""
        return self.__client.get_portfolio_currencies(broker_account_id=self.__account_id) \
            .payload.currencies

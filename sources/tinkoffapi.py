import os
from pprint import pprint

from dotenv import load_dotenv
import tinvest


class TinkoffApi:
    """Обёртка для работы с API Тинькова на основе библиотеки tinvest"""

    def __init__(self):
        self.__init_env_var()

        token = os.getenv('TINKOFF_API_TOKEN')
        self.__client = tinvest.SyncClient(token)

        api = tinvest.PortfolioApi(self.__client)
        account_id = os.getenv("TINKOFF_BROKER_ACCOUNT")
        self.__positions = api.portfolio_get(broker_account_id=account_id) \
            .parse_json().payload.positions
        self.__balance = api.portfolio_currencies_get(broker_account_id=account_id) \
            .parse_json().payload.currencies

    @staticmethod
    def __init_env_var():
        """Инициализация переменных окружения"""
        dotenv_path = os.path.join(os.path.dirname(__file__), '..\.env')
        if os.path.exists(dotenv_path):
            load_dotenv(dotenv_path=dotenv_path)

    def get_usd_course(self):
        """Получение курса USD"""
        return tinvest.MarketApi(self.__client).market_orderbook_get(figi="BBG0013HGFT4", depth=1) \
            .parse_json().payload.last_price

    def get_portfolio_positions_and_balance(self):
        """Получение всех позиций портфеля и баланса в валюте"""
        return self.__positions, self.__balance

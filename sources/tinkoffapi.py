import os
from decimal import Decimal
from datetime import datetime
from pytz import timezone

from dotenv import load_dotenv
import tinvest


class TinkoffApi:
    """Обёртка для работы с API Тинькова на основе библиотеки tinvest"""

    def __init__(self):
        self.__init_env_var()

        self.__client = tinvest.SyncClient(os.getenv('TINKOFF_API_TOKEN'))
        self.__account_id = os.getenv("TINKOFF_BROKER_ACCOUNT")
        self.__broker_account_started_at = datetime.strptime(os.getenv("TINKOFF_ACCOUNT_STARTED"), '%d.%m.%Y')

    @staticmethod
    def __init_env_var():
        """Инициализация переменных окружения"""
        dotenv_path = os.path.join(os.path.dirname(__file__), '../.env')
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

    def get_portfolio_operations(self) -> list:
        """Получение операций, совершенных с портфелем"""
        moscow_tz = timezone('Europe/Moscow')
        from_ = moscow_tz.localize(self.__broker_account_started_at)
        now = moscow_tz.localize(datetime.now())

        return self.__client.get_operations(broker_account_id=self.__account_id, from_=from_, to=now) \
            .payload.operations

    def get_candle_from_date(self, figi, from_, to, interval="15min"):
        """Получение исторических свечей по figi"""
        return self.__client.get_market_candles(figi=figi, from_=from_, to=to,
                                                interval=tinvest.CandleResolution(interval))

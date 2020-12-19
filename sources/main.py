import os
from pprint import pprint

from dotenv import load_dotenv
import tinvest

dotenv_path = os.path.join(os.path.dirname(__file__), '../.env')
if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path=dotenv_path)

TOKEN = os.getenv('TINKOFF_API_TOKEN')
ACCOUNT_ID = os.getenv("TINKOFF_BROKER_ACCOUNT")

client = tinvest.SyncClient(TOKEN)
api = tinvest.PortfolioApi(client)

positions = api.portfolio_get(broker_account_id=ACCOUNT_ID).parse_json().payload.positions
balance = api.portfolio_currencies_get(broker_account_id=ACCOUNT_ID).parse_json().payload.currencies[0]

csv_rows = [';'.join(['Имя', 'Котировка', 'Количество', 'Цена', 'Валюта', 'Заработок'])]
for position in positions:
    csv_rows.append(';'.join(map(str, [position.name,
                                       position.ticker,
                                       position.balance,
                                       position.average_position_price.value,
                                       position.average_position_price.currency.value,
                                       position.expected_yield.value])))
csv_rows.append(';'.join(map(str, ['Баланс', balance.balance])))

with open('../Инвестиции.csv', 'w') as f:
    f.write("\n".join(csv_rows))

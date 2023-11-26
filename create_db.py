import os
import openpyxl
import pandas as pd


from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from json import loads, dumps
from dotenv import load_dotenv
from pathlib import Path
from datetime import datetime

from app import Base, Order


load_dotenv()
env_path = Path('.')/'.env'
load_dotenv(dotenv_path=env_path)

exl_file_path = os.getenv('file_path')
exl_file_name = os.getenv('file_name')
sheet1 = os.getenv('sheet1')
sheet2 = os.getenv('sheet2')
sheet3 = os.getenv('sheet3')
sheet4 = os.getenv('sheet4')
coll_name1 = os.getenv('coll_name1')
search_excpression = os.getenv('search_excpression')

try:
    with pd.ExcelFile(exl_file_path + exl_file_name, engine='openpyxl') as xls:
        df_orders = pd.read_excel(xls, sheet1)
except:
    print('По указанному пути файл не обнаружен')

df_orders = df_orders[df_orders[coll_name1].str.contains(search_excpression)]
df_orders.fillna(0, inplace=True)
df_orders.insert(0, 'Составной ключ', df_orders['Заказ на сборку'].apply(lambda x: x[43:47]) + df_orders['Заказ на сборку'].apply(lambda x: x[29:33]))
df_orders.insert(1, 'Дата заказа', df_orders['Заказ на сборку'].apply(lambda x: x[37:47]))
df_orders.insert(2, 'Номер заказа', df_orders['Заказ на сборку'].apply(lambda x: x[29:33]))

result = df_orders.to_json(orient='records', force_ascii=False)
result_dict = eval(result)

engine = create_engine('sqlite:///data/orders.db')
Base.metadata.drop_all(engine)
Base.metadata.create_all(engine)
session = Session(bind=engine)

for _ in range(len(result_dict)):
    order = Order()
    order.composite_key = result_dict[_]['Составной ключ']
    order.order_date = result_dict[_]['Дата заказа']
    order.order_main = result_dict[_]['Номер заказа']
    order.order_number = result_dict[_]['Заказ на сборку']
    order.nomenclature_name = result_dict[_]['Номенклатура']
    order.oredered = result_dict[_]['Заказано']
    order.released = result_dict[_]['Выпущено']
    order.remains_to_release = result_dict[_]['Осталось выпустить']
    session.add(order)

session.commit()

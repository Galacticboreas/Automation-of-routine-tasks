import os
import openpyxl
import pandas as pd


from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from json import loads, dumps
from dotenv import load_dotenv
from pathlib import Path
from datetime import datetime

from app import Base, Order, ReleaseOfAssemblyKits


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

# Открываем файл с исходными данными по заказам и первый лист
try:
    with pd.ExcelFile(exl_file_path + "/" + exl_file_name, engine='openpyxl') as xls:
        df_orders = pd.read_excel(xls, sheet1)
except:
    print('По указанному пути файл не обнаружен')

# Сортируем по ключевому слову
df_orders = df_orders[df_orders[coll_name1].str.contains(search_excpression)]
df_orders.fillna(0, inplace=True)
# Получаем исходные заказы и добавляем колонки:
# Составной ключ        - колонка №1
# Дата заказа           - колонка №2
# Короткий номер заказа - колонка №3
df_orders.insert(0, 'Составной ключ', df_orders['Заказ на сборку'].apply(lambda x: x[43:47]) + df_orders['Заказ на сборку'].apply(lambda x: x[29:33]))
df_orders.insert(1, 'Дата заказа', df_orders['Заказ на сборку'].apply(lambda x: x[37:47]))
df_orders.insert(2, 'Номер заказа', df_orders['Заказ на сборку'].apply(lambda x: x[29:33]))

# Получаем данные о выпусках подразделений с второго листа
try:
    with pd.ExcelFile(exl_file_path + "/" + exl_file_name, engine='openpyxl') as xls:
        df_kits = pd.read_excel(xls, sheet2)
except:
    print('По указанному пути файл не обнаружен')

# Сортируем по ключевому слову
df_kits = df_kits[df_kits[coll_name1].str.contains(search_excpression)]
df_kits.fillna(0, inplace=True)
# Добавляем колонку с составным ключем
df_kits.insert(0, 'Составной ключ', df_kits['Заказ на сборку'].apply(lambda x: x[43:47]) + df_kits['Заказ на сборку'].apply(lambda x: x[29:33]))
# Удаляем лишние колонки
df_kits = df_kits.drop('Заказ на сборку', axis=1)
df_kits = df_kits.drop('Номенклатура', axis=1)

# Записываем данные в формате json
kits_json = df_kits.to_json(orient='split', force_ascii=False)
orders_json = df_orders.to_json(orient='records', force_ascii=False)

# Преобразуем в словарь
kits_json = eval(kits_json)
kits_json = kits_json['data']
kits_dict = {item[0]: item[1:] for item in kits_json}
orders_json = eval(orders_json)

engine = create_engine('sqlite:///data/orders.db')
Base.metadata.drop_all(engine)
Base.metadata.create_all(engine)
session = Session(bind=engine)

for _ in range(len(orders_json)):
    key = orders_json[_]['Составной ключ']
    order = Order()
    releaseassemblykits = ReleaseOfAssemblyKits()
    order.composite_key = key
    order.order_date = orders_json[_]['Дата заказа']
    order.order_main = orders_json[_]['Номер заказа']
    order.order_number = orders_json[_]['Заказ на сборку']
    order.nomenclature_name = orders_json[_]['Номенклатура']
    order.oredered = orders_json[_]['Заказано']
    order.released = orders_json[_]['Выпущено']
    order.remains_to_release = orders_json[_]['Осталось выпустить']
    if kits_dict.get(key):
        releaseassemblykits.cutting_shop_for_assembly = kits_dict[key][0]
        releaseassemblykits.cutting_shop_for_painting = kits_dict[key][1]
        releaseassemblykits.paint_shop_for_assembly = kits_dict[key][2]
        releaseassemblykits.assembly_shop = kits_dict[key][3]
        releaseassemblykits.order = order
    else:
        releaseassemblykits.cutting_shop_for_assembly = 0
        releaseassemblykits.cutting_shop_for_painting = 0
        releaseassemblykits.paint_shop_for_assembly = 0
        releaseassemblykits.assembly_shop = 0
        releaseassemblykits.order = order
    session.add(order)

session.commit()

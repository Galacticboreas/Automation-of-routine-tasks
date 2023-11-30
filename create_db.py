import os
import openpyxl
import pandas as pd


from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from json import loads, dumps
from dotenv import load_dotenv
from pathlib import Path
from datetime import datetime

from app import Base, Order, ReleaseOfAssemblyKits, \
                DescriptionMainOrder, DescriptionAdditionalOrder


load_dotenv()
env_path = Path('.')/'.env'
load_dotenv(dotenv_path=env_path)

exl_file_path = os.getenv('file_path')
exl_file_name = os.getenv('file_name')
sheet1 = os.getenv('sheet1')
sheet2 = os.getenv('sheet2')
sheet3 = os.getenv('sheet3')
sheet4 = os.getenv('sheet4')
company = os.getenv('company')
coll_name1 = os.getenv('coll_name1')
search_excpression = os.getenv('search_excpression')

# Открываем файл с исходными данными по заказам и первый лист
try:
    with pd.ExcelFile(exl_file_path + "/" + exl_file_name, engine='openpyxl') as xls:
        df_orders = pd.read_excel(xls, sheet1)
        df_kits = pd.read_excel(xls, sheet2)
        df_devision = pd.read_excel(xls, sheet4)
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

# Сортируем по ключевому слову
df_kits = df_kits[df_kits[coll_name1].str.contains(search_excpression)]
df_kits.fillna(0, inplace=True)

# Добавляем колонку с составным ключем
df_kits.insert(0, 'Составной ключ', df_kits['Заказ на сборку'].apply(lambda x: x[43:47]) + df_kits['Заказ на сборку'].apply(lambda x: x[29:33]))

# Удаляем лишние колонки
df_kits = df_kits.drop('Заказ на сборку', axis=1)
df_kits = df_kits.drop('Номенклатура', axis=1)

df_devision = df_devision[df_devision['Организация'].str.contains(company)]
df_devision.fillna(0, inplace=True)

convert_coll = ['Номер', 'Дата', 'Дата запуска', 'Дата исполнения']
df_devision[convert_coll] = df_devision[convert_coll].astype('string')

df_devision['Номер'] = df_devision['Номер'].apply(lambda x: x.zfill(4))

df_devision.insert(0, "Составной ключ", df_devision['Дата'].apply(lambda x: x[6:10]) + df_devision['Номер'])

df_div_main = df_devision[df_devision['Основной заказ на производство'] == 0]
df_div_main = df_div_main.drop(['Номер', 'Организация', 'Основной заказ на производство'], axis=1)

convert_coll = ['Основной заказ на производство']
df_devision[convert_coll] = df_devision[convert_coll].astype('string')
df_div_add = df_devision[df_devision[convert_coll[0]].str.contains(search_excpression)]
df_div_add = df_div_add.drop('Составной ключ', axis=1)
df_div_add.insert(0, 'Составной ключ', df_div_add[convert_coll[0]].apply(lambda x: x[43:47]) + df_div_add[convert_coll[0]].apply(lambda x: x[29:33]))
drop_coll = ['Организация', 'Основной заказ на производство']
df_div_add = df_div_add.drop(drop_coll, axis=1)

# Записываем данные в формате json
kits_json = df_kits.to_json(orient='split', force_ascii=False)
orders_json = df_orders.to_json(orient='records', force_ascii=False)
df_div_main = df_div_main.to_json(orient='split', force_ascii=False)
df_div_add = df_div_add.to_json(orient='split', force_ascii=False)

# Преобразуем в словарь
kits_json = eval(kits_json)
orders_json = eval(orders_json)
df_div_main = eval(df_div_main)
df_div_add = eval(df_div_add)

# Извлекаем данные
df_div_main = df_div_main['data']
df_div_add = df_div_add['data']
kits_json = kits_json['data']

# Преобразуем в словарь
kits_dict = {item[0]: item[1:] for item in kits_json}
df_div_main = {item[0]: item[1:] for item in df_div_main}
df_div_add = {item[0]: item[1:] for item in df_div_add}

# Создаем таблицы в БД
engine = create_engine('sqlite:///data/orders.db')
Base.metadata.drop_all(engine)
Base.metadata.create_all(engine)
session = Session(bind=engine)

# Наполняем таблицы данными из листов Excel
for _ in range(len(orders_json)):
    key = orders_json[_]['Составной ключ']
    order = Order()
    releaseassemblykits = ReleaseOfAssemblyKits()
    descr_main_oreder = DescriptionMainOrder()
    descr_addi_order = DescriptionAdditionalOrder()
    
    # Заполняем таблицу с основным заказом
    order.composite_key = key
    order.order_date = orders_json[_]['Дата заказа']
    order.order_main = orders_json[_]['Номер заказа']
    order.order_number = orders_json[_]['Заказ на сборку']
    order.nomenclature_name = orders_json[_]['Номенклатура']
    order.oredered = orders_json[_]['Заказано']
    order.released = orders_json[_]['Выпущено']
    order.remains_to_release = orders_json[_]['Осталось выпустить']

    # Заполняем таблицу с комплектами для сборки
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
    
    # Заполняем таблицу с подразделением и комментарием по заказу
    if df_div_main.get(key):
        pass
    else:
        pass

    # Заполняем таблицу с описанием дополнительных заказов
    if df_div_add.get(key):
        pass
    else:
        pass
    session.add(order)

session.commit()

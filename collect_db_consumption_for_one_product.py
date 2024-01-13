import os
from collections import defaultdict
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

from app import ArticleExtractor

load_dotenv()
env_path = Path('.')/'.env'
load_dotenv(dotenv_path=env_path)

exl_file_path = os.getenv("file_path")
exl_file_name = os.getenv('file_name_coll_db_cons_per_one_prod')
exl_file_name_orders = os.getenv("file_name_orders")
sheet_main = os.getenv("sheet_main")
sheet_all_cat = os.getenv("sheet_all_categories")
sheet_sheet_material = os.getenv("sheet_category_sheet_material")
forecast = os.getenv("forecast")
not_deployed = os.getenv("not_deployed")
expanded = os.getenv("expanded")
at_work = os.getenv("at_work")
exl_file_dir = os.getenv("exl_file_dir")

# Открываем исходные данные по расходу материалов на заказ
sheet1 = sheet_all_cat + " " + forecast
try:
    with pd.ExcelFile(exl_file_dir + exl_file_path + exl_file_name, engine='openpyxl') as xls:
        df_all_categories = pd.read_excel(xls, sheet1)
except FileNotFoundError:
    print('По указанному пути файл не обнаружен')
df_all_categories.fillna(0, inplace=True)

# Открываем файл с заказами на производство
sheet1 = sheet_main
try:
    with pd.ExcelFile(exl_file_dir + exl_file_path + exl_file_name_orders, engine='openpyxl') as xls:
        df_orders = pd.read_excel(xls, sheet1)
except FileNotFoundError:
    print('По указанному пути файл не обнаружен')
df_orders = df_orders[5:]
df_orders.fillna(0, inplace=True)

# Записываем данные по заказу
orders_data = defaultdict(list)
for index, row in df_orders.iterrows():
    key = row['Заказ (исходный)']
    if not orders_data.get(key):
        if 'Заказ на произв' in key:
            orders_data[key].append([
                row['Заказано'],
                row['Наименование'],
            ])

extractor = ArticleExtractor()

for i in orders_data.keys():
    article = orders_data[i][0][1]
    extractor.get_article(article)

# Заполняем колонку с изделием


def lambdafunc(row):
    if row:
        return row[0][1]
    return 0


df_all_categories['Изделие'] = df_all_categories.apply(lambda row: lambdafunc(orders_data.get(row['Заказ на производство'])), axis=True)

# Извлекаем артикул из наименования изделия


def insert_article(row):
    article = extractor.get_article(str(row['Изделие']))
    return article


df_all_categories.insert(0, 'Артикул изделия', df_all_categories.apply(insert_article, axis=True))

# Заполняем колонку "Заказано" количеством из заказа на производство


def lambdafunc(row):
    if row:
        return row[0][0]
    return 0


df_all_categories.insert(7, 'Заказано', df_all_categories.apply(lambda row: lambdafunc(orders_data.get(row['Заказ на производство'])), axis=True))

# Расчитываем расход материала на 1 изделие мебели и добавляем колонку "Расход на 1 изделие"
df_all_categories['Расход на 1 изделие'] = df_all_categories['Количество'] / (df_all_categories['Заказано'] + .000000001)

# Удаляем лишние колонки
df_all_categories = df_all_categories.drop('N', axis=True)
df_all_categories = df_all_categories.drop('Заказ на производство', axis=True)
df_all_categories = df_all_categories.drop('Заказано', axis=True)
df_all_categories = df_all_categories.drop('Номенклатура до замены', axis=True)
df_all_categories = df_all_categories.drop('Количество', axis=True)
df_all_categories = df_all_categories.drop('Количество до замены', axis=True)

# Записываем результат в файл Excel
df_all_categories.to_excel(exl_file_dir + exl_file_path + 'Расход на 1 изделие.xlsx', index=False)

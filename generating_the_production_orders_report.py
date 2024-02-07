"""Скрипт генерирует отчет о состоянии заказов на производство
"""
from collections import defaultdict
from pprint import pprint

import pandas as pd
from sqlalchemy import create_engine

from app import Bunch

# Обращаемся к БД Исходные данные по заказам на производство
try:
    engine = create_engine('sqlite:///data/orders_row_data.db')
except FileNotFoundError:
    print('По указанному пути файл не обнаружен')

# Открываем таблицу данных основного заказа
try:
    df_main = pd.read_sql("orders_row_data", engine)
except FileExistsError:
    print("Таблица с таким названмем не обнаружена")

# Меняем расположение колонок
cols = ['id', 'furniture_article', 'composite_key', 'full_order_number', 'furniture_name', 'ordered', 'released', 'remains_to_release']
df_main = df_main[cols + [c for c in df_main.columns if c not in cols]]

# Конвертируем данные колонок в строковый тип данных
convert_colls = ['furniture_article', 'composite_key', 'full_order_number', 'furniture_name']
df_main[convert_colls] = df_main[convert_colls].astype('string')
# df_main

# Открываем таблицу с данными о перемещении комплектов мебели
df_release = pd.read_sql("release_of_assemblykits_row_data", engine)

# Удаляем колонку с id перемещения
df_release = df_release.drop('id', axis=1)

# Меняем расположение колонок
cols = ['order_id', 'cutting_shop_for_assembly', 'cutting_shop_for_painting', 'paint_shop_for_assembly', 'assembly_shop']
df_release = df_release[cols + [c for c in df_release.columns if c not in cols]]

# Переименовываем колонку id заказа
df_release.rename(columns={'order_id': 'id'}, inplace=True)
# df_release

# Объединяем таблицы по id заказа на производство
df_marge = pd.merge(df_main, df_release, how='left', on='id')
df_marge.tail(5)

# Открываем таблицу с данными о проценте готовности раскроечных заказов
df_monitor = pd.read_sql("monitor_for_work_centers", engine)

# Удаляем колонку с id процента готовности
df_monitor = df_monitor.drop('id', axis=1)

# Меняем расположение колонок
cols = ['order_id', 'percentage_of_readiness_to_cut', 'number_of_details_plan', 'number_of_details_fact']
df_monitor = df_monitor[cols + [c for c in df_monitor.columns if c not in cols]]

# Переименовываем id заказа
df_monitor.rename(columns={'order_id': 'id'}, inplace=True)

# Объединяем с итоговой таблицей по id заказа
df_marge = pd.merge(df_marge, df_monitor, how='left', on='id')
# df_marge.tail(5)

# Открываем таблицу с описанием основного и дочерних заказов
df_descriptions = pd.read_sql("descriptions_main_orders_row_data", engine)

# Удаляем колонку с id описания заказов
df_descriptions = df_descriptions.drop('id', axis=1)

# Меняем расположение колонок
cols = ['order_id', 'order_date', 'order_number', 'order_company', 'order_division', 'order_launch_date', 'order_execution_date', 'responsible', 'comment']
df_descriptions = df_descriptions[cols + [c for c in df_descriptions.columns if c not in cols]]

# Преименовываем колонку id заказа
df_descriptions.rename(columns={'order_id': 'id'}, inplace=True)

# Объединяем с итоговой таблицей по id заказа
df_marge = pd.merge(df_marge, df_descriptions, how='left', on='id')

# Меняем расроложение колонок
cols = ['id',
        'furniture_article',
        'composite_key',
        'order_date',
        'order_number',
        'full_order_number',
        'furniture_name',
        'ordered',
        'released',
        'remains_to_release',
        'cutting_shop_for_assembly',
        'cutting_shop_for_painting',
        'paint_shop_for_assembly',
        'assembly_shop',
        'percentage_of_readiness_to_cut',
        'number_of_details_plan',
        'number_of_details_fact',
        'order_company',
        'order_division',
        'order_launch_date',
        'order_execution_date',
        'responsible',
        'comment']
df_marge = df_marge[cols + [c for c in df_marge.columns if c not in cols]]
# df_marge.tail(5)

# Открываем таблицу с процентом готовности раскрой для подчиненных заказов
df_sub_orders = pd.read_sql("sub_orders", engine)

# Удаляем колонку id подчиненных заказов
df_sub_orders = df_sub_orders.drop('id', axis=1)

# Меняем расположение колонок
cols = ['order_id', 'composite_key', 'percentage_of_readiness_to_cut', 'number_of_details_plan', 'number_of_details_fact']
df_sub_orders = df_sub_orders[cols + [c for c in df_sub_orders.columns if c not in cols]]

# Переименовываем колонки таблицы подчиненных заказов
df_sub_orders.rename(columns={'order_id': 'id',
                              'composite_key': 'composite_key_sub',
                              'percentage_of_readiness_to_cut': 'percentage_of_readiness_to_cut_sub',
                              'number_of_details_plan': 'number_of_details_plan_sub',
                              'number_of_details_fact': 'number_of_details_fact_sub'}, inplace=True)
# df_sub_orders

# Открываем таблицу с описание подчиненных заказов
df_sub_orders_descriptions = pd.read_sql("sub_orders_report_description", engine)

# Удаляем колонку id подчиненных заказов
df_sub_orders_descriptions = df_sub_orders_descriptions.drop('id', axis=1)

# Меняем расположение колонок
cols = ['order_id',
        'composite_key',
        'order_date',
        'order_number',
        'order_company',
        'order_division',
        'order_launch_date',
        'order_execution_date',
        'responsible',
        'comment']
df_sub_orders_descriptions = df_sub_orders_descriptions[cols + [c for c in df_sub_orders_descriptions.columns if c not in cols]]

# Переименовываем колонки таблицы описания подчиненных заказов
df_sub_orders_descriptions.rename(columns={
    'order_id': 'id',
    'composite_key': 'composite_key_sub_description',
    'order_date': 'order_date_sub_description',
    'order_number': 'order_number_sub_description',
    'order_company': 'order_company_sub_description',
    'order_division': 'order_division_sub_description',
    'order_launch_date': 'order_launch_date_sub_description',
    'order_execution_date': 'order_execution_date_sub_description',
    'responsible': 'responsible_sub_description',
    'comment': 'comment_sub_description'
}, inplace=True)
# df_sub_orders_descriptions

# Сохраняем данные о проценте готовности по раскрою подчиненных заказах в словарь
sub_orders_percentage = defaultdict(list)
for index, row in df_sub_orders.iterrows():
    key = row['id']
    if sub_orders_percentage.get(key):
        sub_orders_percentage[key].append({
            row['composite_key_sub']: {
                'percentage_of_readiness_to_cut_sub': row['percentage_of_readiness_to_cut_sub'],
                'number_of_details_plan_sub': row['number_of_details_plan_sub'],
                'number_of_details_fact_sub': row['number_of_details_fact_sub'],
            }
        })
    else:
        sub_orders_percentage[key].append({
            row['composite_key_sub']: {
                'percentage_of_readiness_to_cut_sub': row['percentage_of_readiness_to_cut_sub'],
                'number_of_details_plan_sub': row['number_of_details_plan_sub'],
                'number_of_details_fact_sub': row['number_of_details_fact_sub'],
            }
        })

# Сохраняем описание подчиненных заказов в словарь
sub_orders_descriptions = defaultdict(list)
for index, row in df_sub_orders_descriptions.iterrows():
    key = row['id']
    if sub_orders_descriptions.get(key):
        sub_orders_descriptions[key].append({
            row['composite_key_sub_description']: {
                'order_date_sub_description': row['order_date_sub_description'],
                'order_number_sub_description': row['order_number_sub_description'],
                'order_company_sub_description': row['order_company_sub_description'],
                'order_division_sub_description': row['order_division_sub_description'],
                'order_launch_date_sub_description': row['order_launch_date_sub_description'],
                'order_execution_date_sub_description': row['order_execution_date_sub_description'],
                'responsible_sub_description': row['responsible_sub_description'],
                'comment_sub_description': row['comment_sub_description'],
            }
        })
    else:
        sub_orders_descriptions[key].append({
            row['composite_key_sub_description']: {
                'order_date_sub_description': row['order_date_sub_description'],
                'order_number_sub_description': row['order_number_sub_description'],
                'order_company_sub_description': row['order_company_sub_description'],
                'order_division_sub_description': row['order_division_sub_description'],
                'order_launch_date_sub_description': row['order_launch_date_sub_description'],
                'order_execution_date_sub_description': row['order_execution_date_sub_description'],
                'responsible_sub_description': row['responsible_sub_description'],
                'comment_sub_description': row['comment_sub_description'],
            }
        })

# for index, row in df_marge.iterrows():
#     print(row)

Division = Bunch # Использовать набор вместо defaultdict

order = Division(cut=Division(cut_to_buffer="раскрой на буфер", cut_to_paint="раскрой на покаску"), paint_to_buffer=Division(paint="покраска на буфер"))

pprint(order)
# >>>
# {'cut': {'cut_to_buffer': 'раскрой на буфер',
#          'cut_to_paint': 'раскрой на покаску'},
#  'paint_to_buffer': {'paint': 'покраска на буфер'}}

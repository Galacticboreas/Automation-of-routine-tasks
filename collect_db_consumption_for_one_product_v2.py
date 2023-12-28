import os
import pandas as pd
import configparser

from pathlib import Path
from collections import defaultdict

config = configparser.ConfigParser()
config.sections()

config.read('settings.ini', encoding='utf8')

exl_file_dir = config['DEFAULT']['Path_dir']
exl_data_dir = config['DEFAULT']['Path_data']
exl_file_name = config['DEFAULT']['File_name']

def fix_nan_to_0(x):
    return 0 if x in ("") else x

def fix_nan_to_str(x):
    return 'не определено' if x in ("") else x

try:
    with pd.ExcelFile(exl_file_dir + exl_data_dir + exl_file_name) as f:
        df = pd.read_excel(f, converters={
            "Изделие": fix_nan_to_str,
            "Номенклатура до замены": fix_nan_to_str,
            "Количество до замены": fix_nan_to_0,
        })
except:
    print('По указанному пути файл не обнаружен')

sheet = config['DEFAULT']['Sheet_main']
exl_file_name = config['DEFAULT']['File_name_orders']

try:
    with pd.ExcelFile(exl_file_dir + exl_data_dir + exl_file_name) as f:
        df_orders = pd.read_excel(f,
                                  sheet_name=sheet,
                                  usecols=['Заказ (исходный)', 'Наименование', 'Заказано']
                                  )
except:
    print('По указанному пути файл не обнаружен')

df_orders = df_orders[5:]

orders_data = defaultdict(list)
for index, row in df_orders.iterrows():
    key = row['Заказ (исходный)']
    if not orders_data.get(key):
        if 'Заказ на произв' in key:
            orders_data[key].append(
                [
                    row['Заказано'],
                    row['Наименование'],
                ]
            )

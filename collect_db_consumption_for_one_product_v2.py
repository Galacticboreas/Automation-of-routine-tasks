import configparser

from openpyxl import load_workbook
from collections import defaultdict

config = configparser.ConfigParser()

config.read('data/settings.ini', encoding='utf8')

exl_file_dir = config['DEFAULT']['Path_dir']
exl_data_dir = config['DEFAULT']['Path_data']
exl_file_name = config['DEFAULT']['File_name_orders1']
sheet_main = config['DEFAULT']['Sheet_main']
exl_file = exl_file_dir + exl_data_dir + exl_file_name

workbook = load_workbook(filename=exl_file, read_only=True, data_only=True)

sheet = workbook[sheet_main]
orders_data = defaultdict(list)
for value in sheet.iter_rows(min_row=7,
                             max_col=7,
                             values_only=True):
    key = value[3]
    if not orders_data.get(key):
        orders_data[key].append(
            [
                value[4],
                value[5]
            ]
        )

exl_file_name = config['DEFAULT']['File_name_source_db_1']
exl_file = exl_file_dir + exl_data_dir + exl_file_name

workbook = load_workbook(filename=exl_file, read_only=True, data_only=True)

sheet = workbook['Сборная прогноз']

for value in sheet.iter_rows(min_row=1,
                             max_row=1,
                             values_only=True):
    print(value)

import configparser
from collections import defaultdict
from dataclasses import dataclass

from openpyxl import Workbook, load_workbook

from app import ArticleExtractor, OrderData

config = configparser.ConfigParser()

config.read('settings.ini', encoding='utf8')

exl_file_dir = config['DEFAULT']['Path_dir']
exl_data_dir = config['DEFAULT']['Path_data']
exl_file_name = config['DEFAULT']['File_name_orders1']
sheet_main = config['DEFAULT']['Sheet_main']
exl_file = exl_file_dir + exl_data_dir + exl_file_name


workbook = load_workbook(
    filename=exl_file,
    read_only=True,
    data_only=True
    )

sheet = workbook[sheet_main]
ordered_data = defaultdict(list)
extractor = ArticleExtractor()

for value in sheet.iter_rows(min_row=7,
                             max_col=7,
                             values_only=True):
    key = value[3]
    name = value[4]
    ordered = value[5]
    article = extractor.get_article(name)
    if not ordered_data.get(key):
        ordered_data[key].append(
            [
                name,
                article,
                ordered
            ]
        )

sheet_name_categ = config['Sheet.name']['Sheet_all_categories']
sheet_name_status = config['Sheet.name']['Forecast']
sheet_name = sheet_name_categ + " " + sheet_name_status

exl_file_name = config['DEFAULT']['File_name_source_db_1']
exl_file = exl_file_dir + exl_data_dir + exl_file_name

workbook = load_workbook(
    filename=exl_file,
    read_only=True,
    data_only=True
    )

sheet = workbook[sheet_name]

materials = []

for value in sheet.iter_rows(min_row=2,
                             values_only=True):
    key = value[1]
    material_code = value[3]
    material_article = value[4]
    material_name = value[5]
    material_amount = value[7] if value[7] else 0

    if ordered_data.get(key):
        furniture_article = extractor.get_article(ordered_data[key][0][1])
        furniture_name = ordered_data[key][0][0]
        consumption_per_order = ordered_data[key][0][2] if ordered_data[key][0][2] else 0
        try:
            consumption_per_1_product = material_amount / consumption_per_order
        except ZeroDivisionError:
            consumption_per_1_product = 0
        material = OrderData(
            furniture_article=furniture_article,
            furniture_name=furniture_name,
            material_code=material_code,
            material_article=material_article,
            material_name=material_name,
            consumption_per_1_product=consumption_per_1_product
        )
        materials.append(material)
len(materials)

workbook = Workbook()

sheet = workbook.active
sheet.title = sheet_name

sheet.append([
    "Артикул",
    "Наименование",
    "Код материала",
    "Артикул материала",
    "Наименование материала",
    "Расход на 1 изделие",
    ])

for material in materials:
    data = [
        material.furniture_article,
        material.furniture_name,
        material.material_code,
        material.material_article,
        material.material_name,
        material.consumption_per_1_product,
    ]
    sheet.append(data)

workbook.save(
    filename=exl_file_dir + exl_data_dir + 'Расход на 1 изделие.xlsx'
    )

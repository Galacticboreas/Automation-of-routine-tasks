import configparser
from collections import defaultdict

from openpyxl import Workbook, load_workbook

from app import ArticleExtractor, OrderData

config = configparser.ConfigParser()

config.read('settings.ini', encoding='utf8')

exl_file_dir = config['DEFAULT']['Path_dir']
exl_data_dir = config['DEFAULT']['Path_data']
file_name_orders1 = config['DEFAULT']['File_name_orders1']
file_name_orders2 = config['DEFAULT']['File_name_orders2']
file_name_source_db_1 = config['DEFAULT']['File_name_source_db_1']
file_name_source_db_2 = config['DEFAULT']['File_name_source_db_2']
sheet_main = config['DEFAULT']['Sheet_main']
forecast = config['Sheet.name']['Forecast']
not_deploy = config['Sheet.name']['Not_deploy']
deploy = config['Sheet.name']['Deploy']
sheet_all_categories = config['Sheet.name']['Sheet_all_categories']
sheet_sheet_material = config['Sheet.name']['Sheet_sheet_material']

# Пути до файлов с заказами
orders_files = [
    file_name_orders1,
    file_name_orders2,
]

# Данные по заказам на производство
ordered_data_file1 = defaultdict(list)
ordered_data_file2 = defaultdict(list)
data_orders = [
    ordered_data_file1,
    ordered_data_file2,
]

extractor = ArticleExtractor()
ordered_data_result = []
for _ in range(len(orders_files)):
    exl_file = exl_file_dir + exl_data_dir + orders_files[_]
    workbook = load_workbook(
        filename=exl_file,
        read_only=True,
        data_only=True
        )

    sheet = workbook[sheet_main]
    ordered_data = data_orders[_]

    for value in sheet.iter_rows(min_row=7,
                                 max_col=7,
                                 values_only=True):
        key = value[3]       # Полный номер заказа на производство
        name = value[4]      # Наименование изделия мебели
        ordered = value[5]   # Количество изделий в заказе
        article = extractor.get_article(name)   # Артикул изделия мебели
        if not ordered_data.get(key):
            ordered_data[key].append(
                [
                    name,
                    article,
                    ordered
                ]
            )
    ordered_data_result.append(ordered_data)

# Названия листов с данными по материалам из выгрузки
sheets = [
    sheet_all_categories + " " + forecast,
    sheet_sheet_material + " " + forecast,
]

material_files = [
    file_name_source_db_1,
    file_name_source_db_2,
]

materials = []   # Собираем данные по расходу материала на 1 изделие

for _ in range(len(material_files)):
    exl_file = exl_file_dir + exl_data_dir + material_files[_]
    workbook = load_workbook(
        filename=exl_file,
        read_only=True,
        data_only=True,
    )
    for i in range(len(sheets)):
        sheet = workbook[sheets[i]]

        for value in sheet.iter_rows(min_row=2, values_only=True):
            key = value[1]                # Полный номер заказа на производство
            material_code = value[3]      # Код материала
            material_article = value[4]   # Артикул материала
            material_name = value[5]      # Наименованме материала
            # # Количество материала в заказе
            material_amount = value[7] if value[7] else 0

            ordered_data = ordered_data_result[_]
            if ordered_data.get(key):
                furniture_name = ordered_data[key][0][0]
                furniture_article = ordered_data[key][0][1]
                cons_temp = ordered_data[key][0][2]
                consumption_per_order = cons_temp if cons_temp else 0
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

# Названия листов с данными по материалам из выгрузки
sheets = [
    sheet_all_categories + " " + not_deploy,
    sheet_all_categories + " " + deploy,
    sheet_sheet_material + " " + not_deploy,
    sheet_sheet_material + " " + deploy,
]

for _ in range(len(material_files)):
    exl_file = exl_file_dir + exl_data_dir + material_files[_]
    workbook = load_workbook(
        filename=exl_file,
        read_only=True,
        data_only=True,
    )
    for i in range(len(sheets)):
        sheet = workbook[sheets[i]]

        for value in sheet.iter_rows(min_row=2, values_only=True):
            key = value[1]                # Полный номер заказа на производство
            material_code = value[3]      # Код материала
            material_article = value[4]   # Артикул материала
            material_name = value[5]      # Наименованме материала
            # # Количество материала в заказе
            material_amount = value[7] if value[7] else 0

            ordered_data = ordered_data_result[_]
            if ordered_data.get(key):
                furniture_name = ordered_data[key][0][0]
                furniture_article = ordered_data[key][0][1]
                cons_temp = ordered_data[key][0][2]
                consumption_per_order = cons_temp if cons_temp else 0
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

# Файл с результатами расчетов расхода на одно изделие
workbook = Workbook()

sheet = workbook.active
sheet.title = 'Расход на 1 изделие'

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

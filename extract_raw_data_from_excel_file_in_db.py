from openpyxl import load_workbook
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
import re
from pprint import pprint

from app import (ArticleExtractor, Base, DescriptionAdditionalOrderRowData,
                 DescriptionMainOrderRowData, OrderRowData,
                 ReleaseOfAssemblyKitsRowData, ReportSettingsOrders,
                 extract_data_moving_sets_of_furniture)

config = ReportSettingsOrders()
extractor = ArticleExtractor()
expression = config.expression
sheet_name = config.sheet_moving_1C
orders_data = dict()
error_log = dict()

if __name__ == '__main__':
    # Открыть файл с исходными данными
    workbook = load_workbook(
        filename=config.source_file,
        read_only=True,
        data_only=True
    )

    # Собираем данные перемещений
    orders_data = extract_data_moving_sets_of_furniture(
        orders_data=orders_data,
        error_log=error_log,
        workbook=workbook,
        sheet=sheet_name,
        expression=expression,
        extractor=extractor,
        config=config
    )

sheet_name = config.sheet_kits_1C
workbook_sheet = workbook[sheet_name]

for value in workbook_sheet.iter_rows(min_row=2, values_only=True):
    full_order_number = value[0]
    composite_key = full_order_number[43:47] + full_order_number[28:33]
    cutting_workshop_for_assembly = value[2]
    cutting_workshop_for_painting = value[3]
    paint_shop_for_assembly = value[4]
    assembly_shop = value[5]
    if (expression in full_order_number):
        if orders_data.get(composite_key):
            orders_data[composite_key]['release_of_assembly_kits'] = {
                'cutting_workshop_for_assembly': cutting_workshop_for_assembly,
                'cutting_workshop_for_painting': cutting_workshop_for_painting,
                'paint_shop_for_assembly': paint_shop_for_assembly,
                'assembly_shop': assembly_shop,
            }
        
sheet_name = config.sheet_percentage_1C
workbook_sheet = workbook[sheet_name]
child_order_data = dict()

for value in workbook_sheet.iter_rows(min_row=2):
    pattern = r'[0-9]{11}[" "]["о"]["т"][" "][0-9]{2}["."][0-9]{2}["."][0-9]{4}[" "][0-9]{2}[":"][0-9]{2}[":"][0-9]{2}'
    if value[0].font.b:
        order_number = value[0].value
        composite_key = order_number[21:25] + order_number[6:11]
        if orders_data.get(composite_key):
            percentage_of_readiness_to_cut = value[2].value
            number_of_details_plan = value[3].value
            number_of_details_fact = value[4].value
            orders_data[composite_key]['job monitor for work centers'] = {
                'percentage_of_readiness_to_cut': percentage_of_readiness_to_cut,
                'number_of_details_plan': number_of_details_plan,
                'number_of_details_fact': number_of_details_fact,
            }
    else:
        child_order = re.search(pattern, value[0].value)
        if child_order:
            order_number = child_order[0]
            key = order_number[21:25] + order_number[6:11]
            percentage_of_readiness_to_cut = value[2].value
            number_of_details_plan = value[3].value
            number_of_details_fact = value[4].value
            child_order_data = {
                key: {
                    'percentage_of_readiness_to_cut': percentage_of_readiness_to_cut,
                    'number_of_details_plan': number_of_details_plan,
                    'number_of_details_fact': number_of_details_fact,
                }
            }
            if orders_data.get(composite_key) and not orders_data[composite_key].get('child_orders'):
                orders_data[composite_key]['child_orders'] = child_order_data
            if orders_data.get(composite_key) and orders_data[composite_key]['child_orders']:
                orders_data[composite_key]['child_orders'][key] = {
                        'percentage_of_readiness_to_cut': percentage_of_readiness_to_cut,
                        'number_of_details_plan': number_of_details_plan,
                        'number_of_details_fact': number_of_details_fact,
                }

pprint(orders_data['202300920'])

# Создаем таблицы в БД
engine = create_engine('sqlite:///data/orders_row_data.db')
Base.metadata.drop_all(engine)
Base.metadata.create_all(engine)
session = Session(bind=engine)

for key in orders_data:
    if not (key == 'ErrorLog'):
        order = OrderRowData()
        order.composite_key = key
        order.full_order_number = orders_data[key]['moving_sets_of_furniture']['full_order_number']
        order.furniture_name = orders_data[key]['moving_sets_of_furniture']['furniture_name']
        order.furniture_article = orders_data[key]['moving_sets_of_furniture']['furniture_article']
        order.ordered = orders_data[key]['moving_sets_of_furniture']['ordered']
        order.released = orders_data[key]['moving_sets_of_furniture']['released']
        order.remains_to_release = orders_data[key]['moving_sets_of_furniture']['remains_to_release']
        if orders_data[key].get("release_of_assembly_kits"):
            releaseassemblykits = ReleaseOfAssemblyKitsRowData()
            releaseassemblykits.cutting_shop_for_assembly = orders_data[key]['release_of_assembly_kits']['cutting_workshop_for_assembly']
            releaseassemblykits.cutting_shop_for_painting = orders_data[key]['release_of_assembly_kits']['cutting_workshop_for_painting']
            releaseassemblykits.paint_shop_for_assembly = orders_data[key]['release_of_assembly_kits']['paint_shop_for_assembly']
            releaseassemblykits.assembly_shop = orders_data[key]['release_of_assembly_kits']['assembly_shop']
            releaseassemblykits.order = order
        else:
            releaseassemblykits = ReleaseOfAssemblyKitsRowData()
            releaseassemblykits.cutting_shop_for_assembly = 0
            releaseassemblykits.cutting_shop_for_painting = 0
            releaseassemblykits.paint_shop_for_assembly = 0
            releaseassemblykits.assembly_shop = 0
            releaseassemblykits.order = order
        session.add(order)
session.commit()

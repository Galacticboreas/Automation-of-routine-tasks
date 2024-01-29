from pprint import pprint

from openpyxl import load_workbook
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import (ArticleExtractor, Base, DescriptionAdditionalOrderRowData,
                 DescriptionMainOrderRowData, OrderRowData,
                 ReleaseOfAssemblyKitsRowData, ReportSettingsOrders,
                 extract_data_job_monitor_for_work_centers,
                 extract_data_moving_sets_of_furniture,
                 extract_data_release_of_assembly_kits)

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

    # Собираем данные по перемещению комплектов мебели
    orders_data = extract_data_moving_sets_of_furniture(
        orders_data=orders_data,
        error_log=error_log,
        workbook=workbook,
        sheet=sheet_name,
        expression=expression,
        extractor=extractor,
        config=config
    )

    # Собираем данные по выпуску комплектов мебели
    orders_data = extract_data_release_of_assembly_kits(
        orders_data=orders_data,
        workbook=workbook,
        sheet=config.sheet_kits_1C,
        expression=config.expression
    )

    # Собираем данные по проценту готовности заказов
    orders_data = extract_data_job_monitor_for_work_centers(
        orders_data=orders_data,
        workbook=workbook,
        sheet=config.sheet_percentage_1C
    )

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

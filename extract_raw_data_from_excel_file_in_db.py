"""
    Скрипт имитирует работу коннектора программы 1С для экспорта
    данных в БД.
    На данном этапе данные из программы 1С, вручную
    копируются в файл excel на отдельные листы, предназначенные
    для каждого вида отчета.
    Скрипт последовательно проходит по листам excel и собирает
    данные по заказу на производство в словарь.
    После сбора всей информации с листов, скрипт импортирует
    данные из словаря в таблицы БД.
"""

from pprint import pprint

from openpyxl import load_workbook
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import (ArticleExtractor, Base, DescriptionAdditionalOrderRowData,
                 DescriptionMainOrderRowData, OrderRowData,
                 ReleaseOfAssemblyKitsRowData, ReportSettingsOrders, JobMonitorForWorkCenters, ChildOrder,
                 extract_data_job_monitor_for_work_centers,
                 extract_data_moving_sets_of_furniture,
                 extract_data_release_of_assembly_kits)

config = ReportSettingsOrders()
extractor = ArticleExtractor()
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
        sheet=config.sheet_moving_1C,
        expression=config.expression,
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

    # Импортируем данные в БД
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
            if orders_data[key].get('job monitor for work centers'):
                monitor = JobMonitorForWorkCenters()
                monitor.percentage_of_readiness_to_cut = orders_data[key]['job monitor for work centers']['percentage_of_readiness_to_cut']
                monitor.number_of_details_plan = orders_data[key]['job monitor for work centers']['number_of_details_plan']
                monitor.number_of_details_fact = orders_data[key]['job monitor for work centers']['number_of_details_fact']
                monitor.order = order
            if orders_data[key].get('child_orders'):
                for key_child in orders_data[key]['child_orders']:
                    child = ChildOrder()
                    child.composite_key = key_child
                    child.percentage_of_readiness_to_cut = orders_data[key]['child_orders'][key_child]['percentage_of_readiness_to_cut']
                    child.number_of_details_plan = orders_data[key]['child_orders'][key_child]['number_of_details_plan']
                    child.number_of_details_fact = orders_data[key]['child_orders'][key_child]['number_of_details_fact']
                    child.order = order
            session.add(order)
    session.commit()

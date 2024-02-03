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

from datetime import datetime
from pprint import pprint

from openpyxl import load_workbook
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import (ArticleExtractor, Base, ReportSettingsOrders,
                 extract_data_job_monitor_for_work_centers,
                 extract_data_moving_sets_of_furniture,
                 extract_data_production_orders_report,
                 extract_data_release_of_assembly_kits,
                 extract_data_to_report_moving_sets_of_furnuture,
                 import_data_to_db_production_orders)

config = ReportSettingsOrders()
extractor = ArticleExtractor()
orders_data = dict()
error_log = dict()

if __name__ == '__main__':
    # Открыть файл с исходными данными
    print("Открываем файл с исходными данными")
    start = datetime.now()
    workbook = load_workbook(
        filename=config.source_file,
        read_only=True,
        data_only=True
    )
    end = datetime.now()
    total_time = (end - start).total_seconds()
    print(f"Время на открытие файла: {total_time} с")

    # Собираем данные по перемещению комплектов мебели
    orders_data = extract_data_to_report_moving_sets_of_furnuture(orders_data=orders_data,
                                                                  workbook=workbook,
                                                                  sheet=config.sheet_moving_1C,
                                                                  expression=config.expression)
    print(len(orders_data))
    pprint(orders_data['202300920'])

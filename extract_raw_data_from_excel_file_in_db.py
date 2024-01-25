from openpyxl import load_workbook
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import (ArticleExtractor, Base, DescriptionAdditionalOrderRowData,
                 DescriptionMainOrderRowData, OrderRowData,
                 ReleaseOfAssemblyKitsRowData, ReportSettingsOrders,
                 extract_data_moving_sets_of_furniture)

config = ReportSettingsOrders()
extractor = ArticleExtractor()
expression = config.expression
orders_data = dict()
error_log = dict()

if __name__ == '__main__':
    # Открыть файл с исходными данными
    workbook = load_workbook(
        filename=config.source_file,
        read_only=True,
        data_only=True
    )
    orders_data = extract_data_moving_sets_of_furniture(
        orders_data=orders_data,
        error_log=error_log,
        workbook=workbook,
        sheet=config.sheet_moving_1C,
        expression=config.expression,
        extractor=extractor,
        config=config
    )

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
        session.add(order)

session.commit()

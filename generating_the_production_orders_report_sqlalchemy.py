"""Скрипт генерирует отчет о состоянии заказов на производство на текущую дату.

Источником данных является БД в которую предварительно импортируют
исходные данные из 1С Производство.

Отчет является вспомогательным источником данных для краткосрочного
планирования потребности к запуску заказов на производство.

Определяет готовность к сборке изделий и формирует задание на сбору.

Определяет возможность к запуску заказов в производство на основании
приоритета и технической готовности к запуску в цех раскрой.
"""

from datetime import datetime

from openpyxl import Workbook
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import (ORDERS_REPORT_COLUMNS_NAME, ORDERS_REPORT_MAIN_SHEET, Base,
                 DescriptionMainOrderRowData, MonitorForWorkCenters,
                 OrderRowData, ReleaseOfAssemblyKitsRowData,
                 ReportSettingsOrders, SubOrder, SubOrderReportDescription,
                 calculate_percentage_of_painting_readiness,
                 determine_if_there_is_a_painting, percentage_of_assembly)

# Обращаемся к БД Исходные данные по заказам на производство
try:
    engine = create_engine('sqlite:///data/orders_row_data.db')
    session = Session(autoflush=False, bind=engine)
    Base.metadata.create_all(bind=engine)
except FileNotFoundError:
    print('По указанному пути файл не обнаружен')

with session as db:
    orders_report = db.query(OrderRowData).all()

workbook = Workbook()
config = ReportSettingsOrders()

dt = datetime.now()
dt = dt.strftime("%d.%m.%Y %Hч%Mм")

sheet = workbook.active
sheet.title = ORDERS_REPORT_MAIN_SHEET
sheet.append(ORDERS_REPORT_COLUMNS_NAME)

product_is_painted = dict()
# Определить статус изделия: крашеное/не крашеное и наличие корпуса
product_is_painted = determine_if_there_is_a_painting(
    orders_report=orders_report,
    product_is_painted=product_is_painted,
    session_db=db,
    table_in_db=SubOrderReportDescription   # Подчиненные заказы
)
# Записываем данные в файл отчета Excel
for order in orders_report:
    furniture_article = order.furniture_article
    ordered = order.ordered
    released = order.released
    paint_to_paint = ""
    cutting_status = ""
    order_division = ""
    order_launch_date = ""
    order_execution_date = ""
    responsible = ""
    comment = ""

    # Описание основного заказа
    description_main = db.query(DescriptionMainOrderRowData).filter(DescriptionMainOrderRowData.order_id == order.id)
    if description_main:
        for d in description_main:
            order_date = d.order_date
            order_number = d.order_number
            order_division = d.order_division
            order_launch_date = d.order_launch_date
            order_execution_date = d.order_execution_date
            responsible = d.responsible
            comment = d.comment
    cutting_shop_for_assembly = 0
    cutting_shop_for_painting = 0
    paint_shop_for_assembly = 0
    assembly_shop = 0

    # Выпуск комплектов мебели (отчет матсеров)
    release_of_assembly = db.query(ReleaseOfAssemblyKitsRowData).filter(ReleaseOfAssemblyKitsRowData.order_id == order.id)
    if release_of_assembly:
        for r in release_of_assembly:
            cutting_shop_for_assembly = r.cutting_shop_for_assembly
            cutting_shop_for_painting = r.cutting_shop_for_painting
            paint_shop_for_assembly = r.paint_shop_for_assembly
            assembly_shop = r.assembly_shop

    # Процент готовности сборка
    percentg_of_assembly = percentage_of_assembly(
        ordered=ordered,
        released=released,
        assembly_shop=assembly_shop
    )

    # Данные монитора рабочих центров для основного заказа
    percentage_of_readiness_to_cut = ""
    number_of_details_plan = 0
    number_of_details_fact = 0

    # Монитор рабочих центров (основной заказ) (готовность раскрой на буфер и раскрой на покраску)
    monitor = db.query(MonitorForWorkCenters).filter(MonitorForWorkCenters.order_id == order.id)
    if monitor:
        for m in monitor:
            percentage_of_readiness_to_cut = m.percentage_of_readiness_to_cut
            number_of_details_plan = m.number_of_details_plan
            number_of_details_fact = m.number_of_details_fact

    # Монитор рабочих центров (подчиненные заказы)
    monitor_sub = db.query(SubOrder).filter(SubOrder.order_id == order.id)

    # Данные для подчиненных заказов (раскрой на буфер)
    composite_key_cut_to_assembly: str = ""
    percentage_of_readiness_to_cut_to_assembly: float = 0
    number_of_details_plan_cut_to_assembly: int = 0
    number_of_details_fact_cut_to_assembly: int = 0
    type_of_movement_cut_to_assembly: str = ""

    # Данные для подичиненных заказов (раскрой на покраску)
    composite_key_cut_to_paint: str = ""
    percentage_of_readiness_to_cut_to_paint: float = 0
    number_of_details_plan_cut_to_paint: int = 0
    number_of_details_fact_cut_to_paint: int = 0
    type_of_movement_cut_to_paint: str = ""
    if monitor_sub:
        for m_sub in monitor_sub:
            if m_sub.type_of_movement == "Раскрой на буфер":
                composite_key_cut_to_assembly = m_sub.composite_key
                percentage_of_readiness_to_cut_to_assembly = m_sub.percentage_of_readiness_to_cut
                number_of_details_plan_cut_to_assembly = m_sub.number_of_details_plan
                number_of_details_fact_cut_to_assembly = m_sub.number_of_details_fact
                type_of_movement_cut_to_assembly = m_sub.type_of_movement
            if m_sub.type_of_movement == "Раскрой на покраску":
                composite_key_cut_to_paint = m_sub.composite_key
                percentage_of_readiness_to_cut_to_paint = m_sub.percentage_of_readiness_to_cut
                number_of_details_plan_cut_to_paint = m_sub.number_of_details_plan
                number_of_details_fact_cut_to_paint = m_sub.number_of_details_fact
                type_of_movement_cut_to_paint = m_sub.type_of_movement

    # Статус изделия: (крашеное/не крашеное) (есть корпус/нет корпуса)
    painted_status = "не определено"
    cutting_status = "не определено"
    if product_is_painted.get(furniture_article):
        painted_status = product_is_painted[furniture_article]['painted_status']
        cutting_status = product_is_painted[furniture_article]['cutting_status']

    # Расчет процента готовности покраски
    percentage_of_readiness_painting = calculate_percentage_of_painting_readiness(
        painted_status=painted_status,
        cutting_shop_for_painting=cutting_shop_for_painting,
        paint_shop_for_assembly=paint_shop_for_assembly,
    )
    # Итоговые данные отчета
    data = [
        furniture_article,
        order.composite_key,
        order_date,
        order_number,
        order.full_order_number,
        order.furniture_name,
        ordered,
        released,
        order.remains_to_release,
        cutting_shop_for_assembly,
        cutting_shop_for_painting,
        paint_shop_for_assembly,
        assembly_shop,
        painted_status,
        cutting_status,
        percentg_of_assembly,
        percentage_of_readiness_to_cut,
        percentage_of_readiness_painting,
        number_of_details_plan,
        number_of_details_fact,
        order_division,
        order_launch_date,
        order_execution_date,
        responsible,
        comment,
        type_of_movement_cut_to_assembly,
        composite_key_cut_to_assembly,
        percentage_of_readiness_to_cut_to_assembly,
        number_of_details_plan_cut_to_assembly,
        number_of_details_fact_cut_to_assembly,
        type_of_movement_cut_to_paint,
        composite_key_cut_to_paint,
        percentage_of_readiness_to_cut_to_paint,
        number_of_details_plan_cut_to_paint,
        number_of_details_fact_cut_to_paint,

    ]
    sheet.append(data)

# Сохранить данные в файл Excel
workbook.save(filename=config.path_dir + config.path_data + config.report_file_name + " от " + dt + config.not_macros)

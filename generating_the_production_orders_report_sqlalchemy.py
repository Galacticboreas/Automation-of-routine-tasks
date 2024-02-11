from datetime import datetime

from openpyxl import Workbook
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from app import (ORDERS_REPORT_COLUMNS_NAME, ORDERS_REPORT_MAIN_SHEET, Base,
                 DescriptionMainOrderRowData, MonitorForWorkCenters,
                 OrderRowData, ReleaseOfAssemblyKitsRowData,
                 ReportSettingsOrders, SubOrder, SubOrderReportDescription)

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

for order in orders_report:
    paint_to_paint = ""
    cutting_status = ""
    description_main = db.query(DescriptionMainOrderRowData).filter(DescriptionMainOrderRowData.order_id == order.id)
    if description_main:
        for d in description_main:
            order_date = d.order_date
            order_number = d.order_number
    cutting_shop_for_assembly = 0
    cutting_shop_for_painting = 0
    paint_shop_for_assembly = 0
    assembly_shop = 0
    release_of_assembly = db.query(ReleaseOfAssemblyKitsRowData).filter(ReleaseOfAssemblyKitsRowData.order_id == order.id)
    if release_of_assembly:
        for r in release_of_assembly:
            cutting_shop_for_assembly = r.cutting_shop_for_assembly
            cutting_shop_for_painting = r.cutting_shop_for_painting
            paint_shop_for_assembly = r.paint_shop_for_assembly
            assembly_shop = r.assembly_shop
    percentage_of_readiness_to_cut = ""
    monitor = db.query(MonitorForWorkCenters).filter(MonitorForWorkCenters.order_id == order.id)
    if monitor:
        for m in monitor:
            percentage_of_readiness_to_cut = m.percentage_of_readiness_to_cut
    data = [
        order.furniture_article,
        order.composite_key,
        order_date,
        order_number,
        order.full_order_number,
        order.furniture_name,
        order.ordered,
        order.released,
        order.remains_to_release,
        cutting_shop_for_assembly,
        cutting_shop_for_painting,
        paint_shop_for_assembly,
        assembly_shop,
        paint_to_paint,
        cutting_status,
        percentage_of_readiness_to_cut,
    ]
    sheet.append(data)

workbook.save(filename=config.path_dir + config.path_data + config.report_file_name + " от " + dt + config.not_macros)

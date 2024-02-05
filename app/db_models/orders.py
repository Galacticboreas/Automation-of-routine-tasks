from dataclasses import dataclass, field
from enum import Enum

from sqlalchemy import Column, Float, ForeignKey, Integer, String
from sqlalchemy.orm import DeclarativeBase, relationship
from tqdm import tqdm

from app import ArticleExtractor


class Report:

    extractor = ArticleExtractor()

    def __init__(self):
        pass


@dataclass
class ReportReleaseOfAssemblyKits():
    cutting_workshop_for_assembly: int = 0
    cutting_workshop_for_painting: int = 0
    paint_shop_for_assembly: int = 0
    assembly_shop: int = 0


@dataclass
class ReportMonitorForWorkCenter():
    percentage_of_readiness_to_cut: float = None
    number_of_details_plan: int = 0
    number_of_details_fact: int = 0


@dataclass
class ReportMovingSetsOfFurniture():
    ordered: int = 0
    released: int = 0
    remains_to_release: int = 0


@dataclass
class ReportDescriptionOfProductionOrder():
    composite_key: str = ""
    order_date: str = ""
    order_number: str = ""
    order_company: str = ""
    order_division: str = ""
    order_launch_date: str = ""
    order_execution_date: str = ""
    responsible: str = ""
    comment: str = ""


@dataclass
class ReportSubOrder():
    report_monitor_for_work_center: dict = field(default_factory=dict)
    report_description_of_production_orders: dict = field(default_factory=dict)


@dataclass
class ReportMainOrder(Report):
    full_order_number: str = None
    furniture_name: str = ""
    furniture_article: str = ""
    report_moving_sets_of_furniture: object = None
    report_release_of_assembly_kits: object = None
    report_monitor_for_work_center: object = None
    report_sub_order: object = None
    report_description_main_order: object = None

    def __post_init__(self):
        self.furniture_article = Report.extractor.get_article(self.furniture_name)


@dataclass
class OrderData:
    furniture_article: str
    furniture_name: str
    material_code: str
    material_article: str
    material_name: str
    consumption_per_1_product: float


class OrderColumn(Enum):
    FULL_ORDER_NUMBER = 3
    FURNITURE_NAME = 4
    ORDERED = 5


class MaterialColumn(Enum):
    FULL_ORDER_NUMBER = 1
    MATERIAL_CODE = 3
    MATERIAL_ARTICLE = 4
    MATERIAL_NAME = 5
    MATERIAL_AMOUNT = 7


class Base(DeclarativeBase):
    pass


class OrderRowData(Base):
    __tablename__ = 'orders_row_data'
    id = Column(Integer, primary_key=True)
    composite_key = Column(String(9), nullable=False)
    full_order_number = Column(String(), nullable=False)
    furniture_name = Column(String())
    furniture_article = Column(String())
    ordered = Column(Integer(), default=0)
    released = Column(Integer(), default=0)
    remains_to_release = Column(Integer(), default=0)
    release_of_assemblykits = relationship("ReleaseOfAssemblyKitsRowData",
                                           uselist=False,
                                           back_populates="order",
                                           cascade="all, delete")
    description_main_order = relationship("DescriptionMainOrderRowData",
                                          uselist=False,
                                          back_populates='order',
                                          cascade="all, delete")
    monitor_for_work_center = relationship("MonitorForWorkCenters",
                                           back_populates='order',
                                           cascade="all, delete",
                                           uselist=False)
    sub_order = relationship("SubOrder", back_populates='order', cascade="all, delete")
    sub_order_report_description = relationship("SubOrderReportDescription",
                                                back_populates='order',
                                                cascade="all, delete")


class ReleaseOfAssemblyKitsRowData(Base):
    __tablename__ = "release_of_assemblykits_row_data"
    id = Column(Integer(), primary_key=True)
    cutting_shop_for_assembly = Column(Integer(), default=0)
    cutting_shop_for_painting = Column(Integer(), default=0)
    paint_shop_for_assembly = Column(Integer(), default=0)
    assembly_shop = Column(Integer(), default=0)
    order_id = Column(Integer(), ForeignKey("orders_row_data.id"))
    order = relationship("OrderRowData", back_populates="release_of_assemblykits")


class MonitorForWorkCenters(Base):
    __tablename__ = 'monitor_for_work_centers'
    id = Column(Integer(), primary_key=True)
    percentage_of_readiness_to_cut = Column(Float(), default=0)
    number_of_details_plan = Column(Integer(), default=0)
    number_of_details_fact = Column(Integer(), default=0)
    order_id = Column(Integer(), ForeignKey("orders_row_data.id"))
    order = relationship("OrderRowData",
                         back_populates="monitor_for_work_center")


class SubOrder(Base):
    __tablename__ = 'sub_orders'
    id = Column(Integer(), primary_key=True)
    composite_key = Column(String(9), nullable=False)
    percentage_of_readiness_to_cut = Column(Float(), default=0)
    number_of_details_plan = Column(Integer(), default=0)
    number_of_details_fact = Column(Integer(), default=0)
    order_id = Column(Integer(), ForeignKey("orders_row_data.id"))
    order = relationship("OrderRowData",
                         back_populates="sub_order")


class SubOrderReportDescription(Base):
    __tablename__ = 'sub_orders_report_description'
    id = Column(Integer(), primary_key=True)
    composite_key = Column(String(9), nullable=False)
    order_date = Column(String())
    order_number = Column(String())
    order_company = Column(String())
    order_division = Column(String())
    order_launch_date = Column(String(), default="")
    order_execution_date = Column(String(), default="")
    responsible = Column(String())
    comment = Column(String(), default="")
    order_id = Column(Integer(), ForeignKey("orders_row_data.id"))
    order = relationship("OrderRowData", back_populates="sub_order_report_description")


class DescriptionMainOrderRowData(Base):
    __tablename__ = 'descriptions_main_orders_row_data'
    id = Column(Integer(), primary_key=True)
    order_date = Column(String())
    order_number = Column(String())
    order_company = Column(String())
    order_division = Column(String())
    order_launch_date = Column(String(), default="")
    order_execution_date = Column(String(), default="")
    responsible = Column(String(), default="")
    comment = Column(String(), default="")
    order_id = Column(Integer(), ForeignKey("orders_row_data.id"))
    order = relationship("OrderRowData", back_populates="description_main_order")


def import_data_to_db_main_orders(orders_data: dict, session: object) -> object:
    for key in tqdm(orders_data, ncols=80, ascii=True, desc='Импорт в БД: основной заказ'):
        # Основной заказ на производство (данные учете)
        order_main = OrderRowData()
        order_main.composite_key = key
        order_main.full_order_number = orders_data[key].full_order_number
        order_main.furniture_name = orders_data[key].furniture_name
        order_main.furniture_article = orders_data[key].furniture_article
        order_main.ordered = orders_data[key].report_moving_sets_of_furniture.ordered
        order_main.released = orders_data[key].report_moving_sets_of_furniture.released
        order_main.remains_to_release = orders_data[key].report_moving_sets_of_furniture.remains_to_release

        # Перемещение комплектов мебели (данные учета)
        if orders_data[key].report_release_of_assembly_kits:
            release = ReleaseOfAssemblyKitsRowData()
            release.cutting_shop_for_assembly = orders_data[key].report_release_of_assembly_kits.cutting_workshop_for_assembly
            release.cutting_shop_for_painting = orders_data[key].report_release_of_assembly_kits.cutting_workshop_for_painting
            release.paint_shop_for_assembly = orders_data[key].report_release_of_assembly_kits.paint_shop_for_assembly
            release.assembly_shop = orders_data[key].report_release_of_assembly_kits.assembly_shop
            release.order = order_main
        if orders_data.get(key) and not orders_data[key].report_release_of_assembly_kits:
            release = ReleaseOfAssemblyKitsRowData()
            release.cutting_shop_for_assembly = 0
            release.cutting_shop_for_painting = 0
            release.paint_shop_for_assembly = 0
            release.assembly_shop = 0
            release.order = order_main

        # Процент готовности заказа (данные монитора рабочих центров)
        if orders_data[key].report_monitor_for_work_center:
            monitor = MonitorForWorkCenters()
            monitor.percentage_of_readiness_to_cut = orders_data[key].report_monitor_for_work_center.percentage_of_readiness_to_cut
            monitor.number_of_details_plan = orders_data[key].report_monitor_for_work_center.number_of_details_plan
            monitor.number_of_details_fact = orders_data[key].report_monitor_for_work_center.number_of_details_fact
            monitor.order = order_main

        # Данные о дополнительных заказах (раскрой на буфер и раскрой на покраску)
        if orders_data[key].report_sub_order and orders_data[key].report_sub_order.report_monitor_for_work_center:
            for key_sub_order in orders_data[key].report_sub_order.report_monitor_for_work_center:
                sub_order = SubOrder()
                sub_order.composite_key = key_sub_order
                sub_order.percentage_of_readiness_to_cut = orders_data[key].report_sub_order.report_monitor_for_work_center[key_sub_order].percentage_of_readiness_to_cut
                sub_order.number_of_details_plan = orders_data[key].report_sub_order.report_monitor_for_work_center[key_sub_order].number_of_details_plan
                sub_order.number_of_details_fact = orders_data[key].report_sub_order.report_monitor_for_work_center[key_sub_order].number_of_details_fact
                sub_order.order = order_main

        # Описание заказа из отчета "Заказы на производство"
        if orders_data[key].report_description_main_order:
            description_main = DescriptionMainOrderRowData()
            description_main.order_date = orders_data[key].report_description_main_order.order_date
            description_main.order_number = orders_data[key].report_description_main_order.order_number
            description_main.order_company = orders_data[key].report_description_main_order.order_company
            description_main.order_division = orders_data[key].report_description_main_order.order_division
            description_main.order_launch_date = orders_data[key].report_description_main_order.order_launch_date
            description_main.order_execution_date = orders_data[key].report_description_main_order.order_execution_date
            description_main.responsible = orders_data[key].report_description_main_order.responsible
            description_main.comment = orders_data[key].report_description_main_order.comment
            description_main.order = order_main

        # Описание дополнительных заказов к основному заказу на производство
        if orders_data[key].report_sub_order and orders_data[key].report_sub_order.report_description_of_production_orders:
            for key_sub_order in orders_data[key].report_sub_order.report_description_of_production_orders:
                sub_description = SubOrderReportDescription()
                sub_description.composite_key = orders_data[key].report_sub_order.report_description_of_production_orders[key_sub_order].composite_key
                sub_description.order_date = orders_data[key].report_sub_order.report_description_of_production_orders[key_sub_order].order_date
                sub_description.order_number = orders_data[key].report_sub_order.report_description_of_production_orders[key_sub_order].order_number
                sub_description.order_company = orders_data[key].report_sub_order.report_description_of_production_orders[key_sub_order].order_company
                sub_description.order_division = orders_data[key].report_sub_order.report_description_of_production_orders[key_sub_order].order_division
                sub_description.order_launch_date = orders_data[key].report_sub_order.report_description_of_production_orders[key_sub_order].order_launch_date
                sub_description.order_execution_date = orders_data[key].report_sub_order.report_description_of_production_orders[key_sub_order].order_execution_date
                sub_description.responsible = orders_data[key].report_sub_order.report_description_of_production_orders[key_sub_order].responsible
                sub_description.comment = orders_data[key].report_sub_order.report_description_of_production_orders[key_sub_order].comment
                sub_description.order = order_main
        session.add(order_main)
    return session


def import_data_to_db_production_orders(orders_data: dict, session: object) -> object:
    """Функция импорта данных из словаря данных о заказаз на производство
    в таблицы БД.

    Args:
        orders_data (dict): [заказы на производство]
        session (object): sqlalchemy session

    Returns:
        session: sqlalchemy session
    """
    for key in tqdm(orders_data, ncols=80, ascii=True, desc='Импорт в БД'):
        if not (key == 'ErrorLog'):
            # Основной заказ на производство
            order = OrderRowData()
            order.composite_key = key
            order.full_order_number = orders_data[key]['moving_sets_of_furniture']['full_order_number']
            order.furniture_name = orders_data[key]['moving_sets_of_furniture']['furniture_name']
            order.furniture_article = orders_data[key]['moving_sets_of_furniture']['furniture_article']
            order.ordered = orders_data[key]['moving_sets_of_furniture']['ordered']
            order.released = orders_data[key]['moving_sets_of_furniture']['released']
            order.remains_to_release = orders_data[key]['moving_sets_of_furniture']['remains_to_release']

            # Перемещение комплектов мебели (данные учета)
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

            # Процент готовности заказа (данные монитора рабочих центров)
            if orders_data[key].get('job monitor for work centers'):
                monitor = MonitorForWorkCenters()
                monitor.percentage_of_readiness_to_cut = orders_data[key]['job monitor for work centers']['percentage_of_readiness_to_cut']
                monitor.number_of_details_plan = orders_data[key]['job monitor for work centers']['number_of_details_plan']
                monitor.number_of_details_fact = orders_data[key]['job monitor for work centers']['number_of_details_fact']
                monitor.order = order

            # Данные о дополнительных заказах (раскрой на буфер и раскрой на прокраску)
            if orders_data[key].get('child_orders') and orders_data[key]['child_orders'].get('job monitor for work centers'):
                for key_child in orders_data[key]['child_orders']['job monitor for work centers']:
                    child = SubOrder()
                    child.composite_key = key_child
                    child.percentage_of_readiness_to_cut = orders_data[key]['child_orders']['job monitor for work centers'][key_child]['percentage_of_readiness_to_cut']
                    child.number_of_details_plan = orders_data[key]['child_orders']['job monitor for work centers'][key_child]['number_of_details_plan']
                    child.number_of_details_fact = orders_data[key]['child_orders']['job monitor for work centers'][key_child]['number_of_details_fact']
                    child.order = order

            # Описание заказа из отчета "Заказы на производство"
            if orders_data[key].get('report description of production orders'):
                description_main = DescriptionMainOrderRowData()
                description_main.order_date = orders_data[key]['report description of production orders']['order_date']
                description_main.order_number = orders_data[key]['report description of production orders']['order_number']
                description_main.order_company = orders_data[key]['report description of production orders']['order_company']
                description_main.order_division = orders_data[key]['report description of production orders']['order_division']
                description_main.order_launch_date = orders_data[key]['report description of production orders']['order_launch_date']
                description_main.order_execution_date = orders_data[key]['report description of production orders']['order_execution_date']
                description_main.responsible = orders_data[key]['report description of production orders']['responsible']
                description_main.comment = orders_data[key]['report description of production orders']['comment']
                description_main.order = order

            # Описание дополнительных заказов к основному заказу на производство
            if orders_data[key].get('child_orders') and orders_data[key]['child_orders'].get('report description of production orders'):
                for key_child_descr in orders_data[key]['child_orders']['report description of production orders']:
                    child_description = SubOrderReportDescription()
                    child_description.composite_key = key_child_descr
                    child_description.order_date = orders_data[key]['child_orders']['report description of production orders'][key_child_descr]['order_date']
                    child_description.order_number = orders_data[key]['child_orders']['report description of production orders'][key_child_descr]['order_number']
                    child_description.order_company = orders_data[key]['child_orders']['report description of production orders'][key_child_descr]['order_company']
                    child_description.order_division = orders_data[key]['child_orders']['report description of production orders'][key_child_descr]['order_division']
                    child_description.order_launch_date = orders_data[key]['child_orders']['report description of production orders'][key_child_descr]['order_launch_date']
                    child_description.order_execution_date = orders_data[key]['child_orders']['report description of production orders'][key_child_descr]['order_execution_date']
                    child_description.responsible = orders_data[key]['child_orders']['report description of production orders'][key_child_descr]['responsible']
                    child_description.comment = orders_data[key]['child_orders']['report description of production orders'][key_child_descr]['comment']
                    child_description.order = order

            session.add(order)
    return session

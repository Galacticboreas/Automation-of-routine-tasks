from dataclasses import dataclass
from enum import Enum

from sqlalchemy import Column, Float, ForeignKey, Integer, String
from sqlalchemy.orm import DeclarativeBase, relationship


class Base(DeclarativeBase):
    pass


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
    descriptions = relationship("DescriptionAdditionalOrderRowData",
                                back_populates='order')
    monitor_for_work_center = relationship("JobMonitorForWorkCenters",
                                           back_populates='order',
                                           cascade="all, delete",
                                           uselist=False)
    child_order = relationship("ChildOrder",
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


class JobMonitorForWorkCenters(Base):
    __tablename__ = 'job_monitor_for_work_centers'
    id = Column(Integer(), primary_key=True)
    percentage_of_readiness_to_cut = Column(Float())
    number_of_details_plan = Column(Integer())
    number_of_details_fact = Column(Integer())
    order_id = Column(Integer(), ForeignKey("orders_row_data.id"))
    order = relationship("OrderRowData",
                         back_populates="monitor_for_work_center")


class ChildOrder(Base):
    __tablename__ = 'child_orders'
    id = Column(Integer(), primary_key=True)
    composite_key = Column(String(9), nullable=False)
    percentage_of_readiness_to_cut = Column(Float())
    number_of_details_plan = Column(Integer())
    number_of_details_fact = Column(Integer())
    order_id = Column(Integer(), ForeignKey("orders_row_data.id"))
    order = relationship("OrderRowData",
                         back_populates="child_order")


class DescriptionMainOrderRowData(Base):
    __tablename__ = 'descriptions_main_orders_row_data'
    id = Column(Integer(), primary_key=True)
    division = Column(String())
    date_launch = Column(String())
    date_execution = Column(String())
    responsible = Column(String())
    comment = Column(String())
    order_id = Column(Integer(), ForeignKey("orders_row_data.id"))
    order = relationship("OrderRowData", back_populates="description_main_order")


class DescriptionAdditionalOrderRowData(Base):
    __tablename__ = "descriptions_row_data"
    id = Column(Integer(), primary_key=True)
    date_create = Column(String())
    order_number = Column(String())
    division = Column(String())
    date_launch = Column(String())
    date_execution = Column(String())
    responsible = Column(String())
    comment = Column(String())
    order_id = Column(Integer(), ForeignKey("orders_row_data.id"))
    order = relationship("OrderRowData", back_populates="descriptions")


def import_data_to_db_production_orders(orders_data: dict, session: object)-> object:
    """Функция импорта данных из словаря данных о заказаз на производство
    в таблицы БД.

    Args:
        orders_data (dict): [заказы на производство]
        session (object): sqlalchemy session

    Returns:
        session: sqlalchemy session
    """
    for key in orders_data:
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
                monitor = JobMonitorForWorkCenters()
                monitor.percentage_of_readiness_to_cut = orders_data[key]['job monitor for work centers']['percentage_of_readiness_to_cut']
                monitor.number_of_details_plan = orders_data[key]['job monitor for work centers']['number_of_details_plan']
                monitor.number_of_details_fact = orders_data[key]['job monitor for work centers']['number_of_details_fact']
                monitor.order = order

            # Данные о дополнительных заказах (раскрой на буфер и раскрой на прокраску)
            if orders_data[key].get('child_orders'):
                for key_child in orders_data[key]['child_orders']:
                    child = ChildOrder()
                    child.composite_key = key_child
                    child.percentage_of_readiness_to_cut = orders_data[key]['child_orders'][key_child]['percentage_of_readiness_to_cut']
                    child.number_of_details_plan = orders_data[key]['child_orders'][key_child]['number_of_details_plan']
                    child.number_of_details_fact = orders_data[key]['child_orders'][key_child]['number_of_details_fact']
                    child.order = order
            session.add(order)
    return session

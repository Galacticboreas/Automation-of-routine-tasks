from dataclasses import dataclass
from datetime import datetime
from enum import Enum

from sqlalchemy import Column, DateTime, ForeignKey, Integer, String, Float
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

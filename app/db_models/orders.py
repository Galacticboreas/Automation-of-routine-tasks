from dataclasses import dataclass
from datetime import datetime
from enum import Enum

from sqlalchemy import Column, DateTime, ForeignKey, Integer, String
from sqlalchemy.orm import DeclarativeBase, relationship


class Base(DeclarativeBase):
    pass


class Order(Base):
    __tablename__ = 'orders'
    id = Column(Integer, primary_key=True)
    composite_key = Column(String(8), nullable=False, index=True)
    order_date = Column(String())
    order_main = Column(String(), nullable=False)
    order_number = Column(String(4), nullable=False)
    nomenclature_name = Column(String())
    oredered = Column(Integer(), default=0)
    released = Column(Integer(), default=0)
    remains_to_release = Column(Integer(), default=0)
    created_on = Column(DateTime(), default=datetime.now)
    updated_on = Column(DateTime(),
                        default=datetime.now,
                        onupdate=datetime.now)
    releaseofassemblykits = relationship("ReleaseOfAssemblyKits",
                                         uselist=False,
                                         back_populates="order",
                                         cascade="all, delete")
    descriptionmainorder = relationship("DescriptionMainOrder",
                                        uselist=False,
                                        back_populates='order',
                                        cascade="all, delete")
    descriptions = relationship("DescriptionAdditionalOrder",
                                back_populates='order')


class ReleaseOfAssemblyKits(Base):
    __tablename__ = "releaseassemblykits"
    id = Column(Integer(), primary_key=True, index=True)
    cutting_shop_for_assembly = Column(Integer())
    cutting_shop_for_painting = Column(Integer())
    paint_shop_for_assembly = Column(Integer())
    assembly_shop = Column(Integer())
    order_id = Column(Integer(), ForeignKey("orders.id"))
    order = relationship("Order", back_populates="releaseofassemblykits")


class DescriptionMainOrder(Base):
    __tablename__ = 'descriptiosnmainorders'
    id = Column(Integer(), primary_key=True, index=True)
    division = Column(String())
    date_launch = Column(String())
    date_execution = Column(String())
    responsible = Column(String())
    comment = Column(String())
    order_id = Column(Integer(), ForeignKey("orders.id"))
    order = relationship("Order", back_populates="descriptionmainorder")


class DescriptionAdditionalOrder(Base):
    __tablename__ = "descriptions"
    id = Column(Integer(), primary_key=True, index=True)
    date_create = Column(String())
    order_number = Column(String())
    division = Column(String())
    date_launch = Column(String())
    date_execution = Column(String())
    responsible = Column(String())
    comment = Column(String())
    order_id = Column(Integer(), ForeignKey("orders.id"))
    order = relationship("Order", back_populates="descriptions")


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

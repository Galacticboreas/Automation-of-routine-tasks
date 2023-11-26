from sqlalchemy import Integer, String, Column, DateTime, ForeignKey
from sqlalchemy.orm import DeclarativeBase
from datetime import datetime


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


class ReleaseAssemblyKits(Base):
    __tablename__ = 'releaseofassemblykits'
    id = Column(Integer(), primary_key=True)
    cutting_shop_for_assembly = Column(Integer())
    cutting_shop_for_painting = Column(Integer())
    paint_shop_for_assembly = Column(Integer())
    assembly_shop = Column(Integer())
    orders_id = Column(Integer(), ForeignKey('orders.id'), index=True)

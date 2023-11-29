from sqlalchemy import Integer, String, Column, DateTime, ForeignKey
from sqlalchemy.orm import DeclarativeBase, relationship
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
    updated_on = Column(DateTime(), default=datetime.now, onupdate=datetime.now)
    releaseofassemblykits = relationship("ReleaseOfAssemblyKits",
                                         uselist=False,
                                         back_populates="order",
                                         cascade="all, delete")
    # descriptionmainorder = relationship("DescriptionMainOrder",
    #                                     uselist=False,
    #                                     back_populates='order',
    #                                     cascade="all, delete")
    # descriptionadditionalorder = relationship("DescriptionAdditionalOrder",
    #                                           back_populates="order",
    #                                           cascade="all, delete")


class ReleaseOfAssemblyKits(Base):
    __tablename__ = "releaseassemblykits"
    id = Column(Integer(), primary_key=True, index=True)
    cutting_shop_for_assembly = Column(Integer())
    cutting_shop_for_painting = Column(Integer())
    paint_shop_for_assembly = Column(Integer())
    assembly_shop = Column(Integer())
    order_id = Column(Integer(), ForeignKey("orders.id"))
    order = relationship("Order", back_populates="releaseofassemblykits")


# class DescriptionMainOrder(Base):
#     __tablename__ = 'descriptiosnmainorders'
#     id = Column(Integer(), primary_key=True, index=True)
#     organization = Column(String())
#     division = Column(String())
#     date_launch = Column(String())
#     date_execution = Column(String())
#     responsible = Column(String())
#     comment = Column(String())
#     order_id = Column(Integer(), ForeignKey("orders.id"))
#     order = relationship("Order", back_populates="descriptionmainorder")


# class DescriptionAdditionalOrder(Base):
#     __tablename__ = "descriptionsadditionalorders"
#     id = Column(Integer(), primary_key=True, index=True)
#     organization = Column(String())
#     division = Column(String())
#     date_launch = Column(String())
#     date_execution = Column(String())
#     responsible = Column(String())
#     comment = Column(String())
#     order_id = Column(Integer(), ForeignKey("orders.id"))
#     order = relationship("Order", back_populates="orders")

from sqlalchemy import Integer, String, Column, DateTime
from sqlalchemy.orm import DeclarativeBase
from datetime import datetime

class Base(DeclarativeBase): pass

class Order(Base):
    __tablename__= 'order'
    id = Column(Integer, primary_key=True)
    composite_key = Column(String(8), nullable=False, index=True)
    order_date = Column(DateTime())
    order_main = Column(String(), nullable=False)
    order_number = Column(String(4), nullable=False)
    nomenclature_name = Column(String())
    oredered = Column(Integer(), default=0)
    released = Column(Integer(), default=0)
    remains_to_release = Column(Integer(), default=0)
    created_on = Column(DateTime(), default=datetime.now)
    updated_on = Column(DateTime(), default=datetime.now, onupdate=datetime.now)

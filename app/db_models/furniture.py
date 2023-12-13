from sqlalchemy.orm import DeclarativeBase
from sqlalchemy import Integer, DateTime, String, Column

from datetime import datetime


class Base(DeclarativeBase): pass


class Furniture(Base):
    __tablename__= 'furniture'
    id = Column(String(), primary_key=True)
    main_name = Column(String())
    counterparty = Column(String(), nullable=False, default="не определен")
    created_on = Column(DateTime(), default=datetime.now)
    update_on = Column(DateTime(), default=datetime.now, onupdate=datetime.now)

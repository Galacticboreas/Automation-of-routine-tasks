from abc import abstractmethod
from enum import Enum


class Column(Enum):
    FULL_ORDER_NUMBER = 3
    FURNITURE_NAME = 4
    ORDERED = 5

print(Column.ORDERED.name)

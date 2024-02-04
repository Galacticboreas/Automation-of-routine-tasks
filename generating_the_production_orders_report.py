"""Скрипт генерирует отчет о состоянии заказов на производство
"""
import sqlite3
from sqlalchemy import create_engine
from sqlalchemy.orm import Session


# engine = create_engine('sqlite:///data/orders_row_data.db')
# session = Session(bind=engine)

sql = "SELECT * FROM orders_row_data"

with sqlite3.connect('data/orders_row_data.db') as con:
    cursor = con.cursor()
    result = cursor.execute(sql).fetchall()

for key in result:
    print(key[0])

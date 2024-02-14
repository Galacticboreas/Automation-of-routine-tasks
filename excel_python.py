import os
from pprint import pprint
from app import Base, ReportSettingsOrders, extract_data_contractors, import_data_to_db_contractors
from openpyxl import load_workbook
from sqlalchemy.orm import Session
from sqlalchemy import create_engine

config = ReportSettingsOrders()

path_to_file = config.path_dir + config.path_data + config.file_name_orders2 + config.macros

wb = load_workbook(path_to_file, read_only=True, keep_vba=True, data_only=True, keep_links=False)

sheet_name = config.sheet_contractors
workbook = wb[sheet_name]
contractors = dict()
contractors = extract_data_contractors(contractors=contractors,
                                       workbook=workbook,
                                       config=config)
engine = create_engine('sqlite:///data/contractors.db')
Base.metadata.drop_all(engine)
Base.metadata.create_all(engine)
session = Session(bind=engine)
session = import_data_to_db_contractors(contractors=contractors,
                                        session=session)
session.commit()

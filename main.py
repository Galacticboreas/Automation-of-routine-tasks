import os
import pandas as pd
import openpyxl
import json as js
from json import loads, dumps

from dotenv import load_dotenv
from pathlib import Path

load_dotenv()
env_path = Path('.')/'.env'
load_dotenv(dotenv_path=env_path)

excel_file_path = os.getenv('file_path')
excel_file_name = os.getenv('file_name')
sheet1 = os.getenv('sheet1')
sheet2 = os.getenv('sheet2')
sheet3 = os.getenv('sheet3')
sheet4 = os.getenv('sheet4')
coll_name1 = os.getenv('coll_name1')
search_excpression = os.getenv('search_excpression')

if __name__ == '__main__':
    with pd.ExcelFile(excel_file_path + excel_file_name) as xls:
        df_main = pd.read_excel(xls, sheet1)
df_main = df_main[df_main[coll_name1].str.contains(search_excpression)]
df_main.fillna(0, inplace=True)

with pd.ExcelFile(excel_file_path + excel_file_name) as xls:
        df1 = pd.read_excel(xls, sheet2)
df1 = df1[df1[coll_name1].str.contains(search_excpression)]
df1.fillna(0, inplace=True)

df_merge = pd.merge(df_main, df1, how='left')
df_merge.fillna(0, inplace=True)

result = df_merge.to_json(orient='table')
parser = loads(result)
result_end = dumps(parser, indent=4, ensure_ascii=False)

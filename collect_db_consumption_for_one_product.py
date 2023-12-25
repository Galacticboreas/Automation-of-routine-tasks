import os
import pandas as pd
import openpyxl as opxl

from pathlib import Path
from dotenv import load_dotenv
from collections import defaultdict


load_dotenv()
env_path = Path('.')/'.env'
load_dotenv(dotenv_path=env_path)

exl_file_path = os.getenv("file_path")
exl_file_name = os.getenv('file_name_coll_db_cons_per_one_prod')
exl_file_name_orders = os.getenv("file_name_orders")
sheet_main = os.getenv("sheet_main")
sheet_all_cat = os.getenv("sheet_all_categories")
sheet_sheet_material = os.getenv("sheet_category_sheet_material")
forecast = os.getenv("forecast")
not_deployed = os.getenv("not_deployed")
expanded = os.getenv("expanded")
at_work = os.getenv("at_work")
exl_file_dir = os.getenv("exl_file_dir")

# Открываем исходные данные по расходу материалов на заказ
sheet1 = sheet_all_cat + " " + forecast
try:
    with pd.ExcelFile(exl_file_dir + exl_file_path + exl_file_name, engine='openpyxl') as xls:
        df_all_categories = pd.read_excel(xls, sheet1)
except:
    print('По указанному пути файл не обнаружен')
df_all_categories.fillna(0, inplace=True)

# Открываем файл с заказами на производство
sheet1 = sheet_main
try:
    with pd.ExcelFile(exl_file_dir + exl_file_path + exl_file_name_orders, engine='openpyxl') as xls:
        df_orders = pd.read_excel(xls, sheet1)
except:
    print('По указанному пути файл не обнаружен')
df_orders = df_orders[5:]
df_orders.fillna(0, inplace=True)

# Записываем данные по заказу 
orders_data = defaultdict(list)
for index, row in df_orders.iterrows():
    key = row['Заказ (исходный)']
    if not orders_data.get(key):
        if 'Заказ на произв' in key:
            orders_data[key].append([
                row['Заказано'],
                row['Наименование'],
            ])


class ArticleExtractor:

    articles = dict()
    num_of_iter = 0

    def get_article(self, name):
        self.num_of_iter += 1
        if self.articles.get(name):
            return self.articles[name]
        return self.__extract_atricle(name)
    
    def __extract_atricle(self, name):
        num_of_occur = name.count(")")
        temp_name = name
        for i in range(num_of_occur):
            if self.get_num_char(temp_name) == 6:
                art_temp = self.get_symbol(temp_name)
                if art_temp.isdigit():
                    self.articles[temp_name] = art_temp
                    return art_temp
                else:
                    num_right = temp_name.rfind(")")
                    num_left = temp_name.rfind("(")
                    temp_name = temp_name[:num_right] + temp_name[num_right + 1:]
                    temp_name = temp_name[:num_left] + temp_name[num_left + 1:]
            else:
                num_right = temp_name.rfind(")")
                num_left = temp_name.rfind("(")
                temp_name = temp_name[:num_right] + temp_name[num_right + 1:]
                temp_name = temp_name[:num_left] + temp_name[num_left + 1:]
        self.articles[name] = name
        return name
    
    def get_num_char(self, name):
        num_right = name.rfind(")")
        num_left = name.rfind("(") + 1
        return num_right - num_left
    
    def get_symbol(self, name):
        return name[name.rfind("(") + 1: name.rfind(")")]

    def get_num_extract_articles(self):
        return len(self.articles)
    
    def get_num_of_iter(self):
        return self.num_of_iter

extractor = ArticleExtractor()

for i in orders_data.keys():
    article = orders_data[i][0][1]
    extractor.get_article(article)

print(extractor.get_num_extract_articles())
print(extractor.articles)
print(extractor.get_num_of_iter())

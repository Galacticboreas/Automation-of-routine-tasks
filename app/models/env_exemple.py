import os

from pathlib import Path
from dotenv import load_dotenv
from dataclasses import dataclass


@dataclass
class EnvExemple():
    exl_file_path: str
    exl_file_name: str
    sheet1: str
    sheet2: str
    sheet3: str
    sheet4: str
    coll_name: str
    search_excpr: str
    env_path: str

    def load_env_path(self):
        load_dotenv()
        self.env_path = Path('.')/'.env'
        load_dotenv(dotenv_path=self.env_path)

    def get_file_path(self):
        self.exl_file_path = os.getenv('file_path')
        self.exl_file_name = os.getenv('file_name')
        return self.exl_file_path + self.exl_file_name
    
    def get_sheet_name(self, sheet_num):
        if sheet_num == 1:
            self.sheet1 = os.getenv('sheet1')
        elif sheet_num == 2:
            self.sheet2 = os.getenv('sheet2')
        elif sheet_num == 3:
            self.sheet3 = os.getenv('sheet3')
        elif sheet_num == 4:
            self.sheet4 = os.getenv('sheet4')

    def get_coll_name1(self):
        self.coll_name = os.getenv('coll_name1')

    def get_search_excpr(self):
        self.search_excpr = os.getenv('search_excpression')

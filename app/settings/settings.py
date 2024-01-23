from dataclasses import dataclass
import configparser

config = configparser.ConfigParser()

@dataclass
class ReportSettingsOrders:
     exl_file: str = ""

     path_dir: str = ""
     path_data: str = ""
     source_file_name: str = ""
     macros: str = ""
     not_macros: str = ""
     sheet_main: str = ""

     def __init__(self):
          config = configparser.ConfigParser()
          config.read('settings.ini', encoding='utf8')
          self.path_dir = config['PATH.DIR']['Path_dir']
          self.path_data = config["PATH.DIR"]['Path_data']
          self.source_file_name = config["File.name"]['File_name_source']
          self.macros = config['File.extension']['macros']
          self.not_macros = config['File.extension']['not_macros']


     def __post_init(self):
          self.exl_file = self.path_dir + self.path_data + self.source_file_name + self.macros

ord_settings = ReportSettingsOrders()

print("-------------------")
print(ord_settings.exl_file)
print("-------------------")

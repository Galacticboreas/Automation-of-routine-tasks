import configparser
from dataclasses import dataclass, field

@dataclass
class ReportSettingsOrders:
     config = configparser.ConfigParser()
     config.read('settings.ini', encoding='utf8')

     path_dir: str = config['PATH.DIR']['Path_dir']
     path_data: str = config["PATH.DIR"]['Path_data']
     source_file_name: str = config["File.name"]['File_name_source']
     report_file_name: str = config['File.name']['File_name_report']
     macros: str = config['File.extension']['macros']
     not_macros: str = config['File.extension']['not_macros']
     source_file: str = field(init=False)
     report_file: str = field(init=False)

     def __post_init__(self):
          self.source_file = self.path_dir + self.path_data + self.source_file_name + self.not_macros
          self.report_file = self.path_dir + self.path_data + self.report_file_name + self.macros

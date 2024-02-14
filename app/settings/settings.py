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
    sheet_moving_1C: str = config['Sheet.name']['Sheet_moving_1C']
    sheet_kits_1C: str = config['Sheet.name']['Sheet_kits_1C']
    sheet_percentage_1C: str = config['Sheet.name']['Sheet_percentage_1C']
    sheet_division_1C: str = config['Sheet.name']['Sheet_division_1C']
    sheet_monitor_1C: str = config['Sheet.name']['Sheet_monitor_1C']
    sheet_contractors: str = config['Sheet.name']['Sheet_contractors']
    macros: str = config['File.extension']['macros']
    not_macros: str = config['File.extension']['not_macros']
    expression: str = config['Expressions']['Production_order']
    file_name_orders2: str = config['File.name']['File_name_orders2']

    source_file: str = field(init=False)
    report_file: str = field(init=False)

    def __post_init__(self):
        self.source_file = "".join(
            [
               self.path_dir,
               self.path_data,
               self.source_file_name,
               self.not_macros,
            ]
        )
        self.report_file = "".join(
            [
               self.path_dir,
               self.path_data,
               self.report_file_name,
               self.macros,
            ]
        )

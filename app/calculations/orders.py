import re

from tqdm import tqdm

from app.db_models.orders import (Report, ReportDescriptionOfProductionOrder,
                                  ReportMainOrder, ReportMonitorForWorkCenter,
                                  ReportMovingSetsOfFurniture,
                                  ReportReleaseOfAssemblyKits, ReportSubOrder)


def extract_data_to_report_moving_sets_of_furnuture(orders_data: dict,
                                                    workbook: object,
                                                    contractors: dict,
                                                    sheet: str,
                                                    expression: str,) -> dict:
    workbook_sheet = workbook[sheet]

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True),
                      ncols=80, ascii=True, desc=sheet):
        full_order_number = value[0]
        composite_key = full_order_number[43:47] + full_order_number[28:33]
        furniture_name = value[1]
        furniture_article = Report.extractor.get_article(furniture_name)
        furniture_contractor = contractors[furniture_article] if contractors.get(furniture_article) else ""
        ordered = value[2]
        released = value[3]
        remains_to_release = value[4]
        if (expression in full_order_number) and (ordered is not None):
            # Заполняем данные основного заказа
            order_main = ReportMainOrder()
            order_main.full_order_number = full_order_number
            order_main.furniture_name = furniture_name
            order_main.furniture_article = furniture_article
            order_main.furniture_contractor = furniture_contractor
            # Заполняем отчет перемещение комплектов мебели
            report_moving = ReportMovingSetsOfFurniture()
            report_moving.ordered = ordered
            report_moving.released = released
            report_moving.remains_to_release = remains_to_release
            order_main.report_moving_sets_of_furniture = report_moving
            orders_data[composite_key] = order_main
    return orders_data


def extract_data_moving_sets_of_furniture(orders_data: dict,
                                          error_log: dict,
                                          workbook: object,
                                          sheet: str,
                                          expression: str,
                                          extractor: object,
                                          config: object) -> dict:
    """Функция предназначена для сбора данных по перемещению комплектов мебели
    из файла excel в словарь данных.

    Args:
        orders_data (dict): [данные по заказах на производство]
        error_log (dict): [журнал ошибок учета]
        workbook (object): [файл excel с исходными данными из 1С]
        sheet (str): [наименование листа]
        expression (str): [ключевое слово для сортировки строк]
        extractor (object): [функция извлечения 6-ти значного артикула из наименования]
        config (object): [класс с настройками]

    Returns:
        orders_data (dict): [данные по заказам на производство]
    """
    workbook_sheet = workbook[config.sheet_moving_1C]

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True),
                      ncols=80, ascii=True, desc=config.sheet_moving_1C):
        full_order_number = value[0]
        composite_key = full_order_number[43:47] + full_order_number[28:33]
        furniture_name = value[1]
        furniture_article = extractor.get_article(furniture_name)
        ordered = value[2]
        released = value[3]
        remains_to_release = value[4]
        if (expression in full_order_number) and (ordered is not None):
            orders_data[composite_key] = {
                'moving_sets_of_furniture': {
                    'full_order_number': full_order_number,
                    'furniture_name': furniture_name,
                    'furniture_article': furniture_article,
                    'ordered': ordered,
                    'released': released,
                    'remains_to_release': remains_to_release,
                }
                }
        elif (expression in full_order_number) and (ordered is None):
            error_log[composite_key] = {
                'error description': {
                    'type of error': 'ошибка учета',
                    'name of the report': sheet,
                    'troubleshooting steps': 'отправить номер заказа в отдел \
                        учета для корректировки перемещения',
                    },
                'order parameters': {
                    'full_order_number': full_order_number,
                    'furniture_name': furniture_name,
                    'furniture_article': furniture_article,
                    'ordered': ordered,
                    'released': released,
                    'remains_to_release': remains_to_release,
                    }
                    }
    orders_data['ErrorLog'] = error_log
    return orders_data


def extract_data_to_report_release_of_assembly_kits(orders_data: dict,
                                                    workbook: object,
                                                    sheet: str,
                                                    expression: str) -> dict:
    workbook_sheet = workbook[sheet]

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True),
                      ncols=80, ascii=True, desc=sheet):
        full_order_number = value[0]
        composite_key = full_order_number[43:47] + full_order_number[28:33]
        cutting_workshop_for_assembly = value[2]
        cutting_workshop_for_painting = value[3]
        paint_shop_for_assembly = value[4]
        assembly_shop = value[5]
        if (expression in full_order_number):
            if orders_data.get(composite_key):
                report = ReportReleaseOfAssemblyKits()
                report.cutting_workshop_for_assembly = cutting_workshop_for_assembly
                report.cutting_workshop_for_painting = cutting_workshop_for_painting
                report.paint_shop_for_assembly = paint_shop_for_assembly
                report.assembly_shop = assembly_shop
                orders_data[composite_key].report_release_of_assembly_kits = report
    return orders_data


def extract_data_release_of_assembly_kits(orders_data: dict,
                                          workbook: object,
                                          sheet: str,
                                          expression: str,
                                          config: object) -> dict:
    """Функция предназначена для сбора данных из отчета "Перемещение
    комплектов мебели"

    Args:
        orders_data (dict): [данные по заказам на производство]
        workbook (object): [файл excel с исходными данными]
        sheet (str): [наименование листа]
        expression (str): [ключевое слово для сортировки строк]

    Returns:
        orders_data (dict): [данные по заказам на производство]
    """
    workbook_sheet = workbook[sheet]

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True),
                      ncols=80, ascii=True, desc=config.sheet_kits_1C):
        full_order_number = value[0]
        composite_key = full_order_number[43:47] + full_order_number[28:33]
        cutting_workshop_for_assembly = value[2]
        cutting_workshop_for_painting = value[3]
        paint_shop_for_assembly = value[4]
        assembly_shop = value[5]
        if (expression in full_order_number):
            if orders_data.get(composite_key):
                orders_data[composite_key]['release_of_assembly_kits'] = {
                    'cutting_workshop_for_assembly': cutting_workshop_for_assembly,
                    'cutting_workshop_for_painting': cutting_workshop_for_painting,
                    'paint_shop_for_assembly': paint_shop_for_assembly,
                    'assembly_shop': assembly_shop,
                }
    return orders_data


def extract_data_to_report_monitor_for_work_centers(orders_data: dict,
                                                    workbook: object,
                                                    sheet: str) -> dict:
    workbook_sheet = workbook[sheet]

    for value in tqdm(workbook_sheet.iter_rows(min_row=2),
                      ncols=80, ascii=True, desc=sheet):
        pattern = r'[0-9]{11}[" "]["о"]["т"][" "][0-9]{2}["."][0-9]{2}["."][0-9]{4}[" "][0-9]{1,2}[":"][0-9]{2}[":"][0-9]{2}'
        if value[0].font.b:
            order_number = value[0].value
            composite_key = order_number[21:25] + order_number[6:11]
            if orders_data.get(composite_key):
                percentage_of_readiness_to_cut = value[2].value
                number_of_details_plan = value[3].value
                number_of_details_fact = value[4].value
                report_monitor = ReportMonitorForWorkCenter()
                report_monitor.percentage_of_readiness_to_cut = percentage_of_readiness_to_cut
                report_monitor.number_of_details_plan = number_of_details_plan
                report_monitor.number_of_details_fact = number_of_details_fact
                report_monitor.type_of_movement = "Готовность раскрой"
                orders_data[composite_key].report_monitor_for_work_center = report_monitor
        else:
            sub_order = re.search(pattern, value[0].value)
            if sub_order:
                order_number = sub_order[0]
                key = order_number[21:25] + order_number[6:11]
                percentage_of_readiness_to_cut = value[2].value
                number_of_details_plan = value[3].value
                number_of_details_fact = value[4].value

                if orders_data.get(composite_key) \
                        and not orders_data[composite_key].report_sub_order:
                    # Заполняем процент готовности подзаказа
                    report_monitor = ReportMonitorForWorkCenter()
                    report_monitor.percentage_of_readiness_to_cut = percentage_of_readiness_to_cut
                    report_monitor.number_of_details_plan = number_of_details_plan
                    report_monitor.number_of_details_fact = number_of_details_fact
                    report_monitor.type_of_movement = "Раскрой на буфер"
                    # Заполняем отчет для подзаказа по ключу
                    report_sub_order = ReportSubOrder()
                    orders_data[composite_key].report_sub_order = report_sub_order
                    orders_data[composite_key].report_sub_order.report_monitor_for_work_center[key] = report_monitor

                if orders_data.get(composite_key) \
                        and not orders_data[composite_key].report_sub_order.report_monitor_for_work_center.get(key):
                    # Заполняем процент готовности подзаказа
                    report_monitor = ReportMonitorForWorkCenter()
                    report_monitor.percentage_of_readiness_to_cut = percentage_of_readiness_to_cut
                    report_monitor.number_of_details_plan = number_of_details_plan
                    report_monitor.number_of_details_fact = number_of_details_fact
                    report_monitor.type_of_movement = "Раскрой на покраску"
                    # Заполняем отчет для подзаказа по ключу
                    orders_data[composite_key].report_sub_order.report_monitor_for_work_center[key] = report_monitor
    return orders_data


def extract_data_job_monitor_for_work_centers(orders_data: dict,
                                              workbook: object,
                                              sheet: str,
                                              config: object) -> dict:
    """Функция предназначена для сбора данных из отчета "Монитор
    рабочих центров"

    Args:
        orders_data (dict): [данные по заказам на производство]
        workbook (object): [файл excel с исходными данными]
        sheet (str): [наименование листа]

    Returns:
        orders_data (dict): [данные по заказам на производство]
    """
    workbook_sheet = workbook[sheet]

    for value in tqdm(workbook_sheet.iter_rows(min_row=2),
                      ncols=80, ascii=True, desc=config.sheet_percentage_1C):
        pattern = r'[0-9]{11}[" "]["о"]["т"][" "][0-9]{2}["."][0-9]{2}["."][0-9]{4}[" "][0-9]{1,2}[":"][0-9]{2}[":"][0-9]{2}'
        if value[0].font.b:
            order_number = value[0].value
            composite_key = order_number[21:25] + order_number[6:11]
            if orders_data.get(composite_key):
                percentage_of_readiness_to_cut = value[2].value
                number_of_details_plan = value[3].value
                number_of_details_fact = value[4].value
                orders_data[composite_key]['job monitor for work centers'] = {
                    'percentage_of_readiness_to_cut': percentage_of_readiness_to_cut,
                    'number_of_details_plan': number_of_details_plan,
                    'number_of_details_fact': number_of_details_fact,
                }
        else:
            child_order = re.search(pattern, value[0].value)
            if child_order:
                order_number = child_order[0]
                key = order_number[21:25] + order_number[6:11]
                percentage_of_readiness_to_cut = value[2].value
                number_of_details_plan = value[3].value
                number_of_details_fact = value[4].value

                if orders_data.get(composite_key) \
                        and not orders_data[composite_key].get('child_orders'):
                    orders_data[composite_key]['child_orders'] = {
                        'job monitor for work centers': {
                            key: {
                                'percentage_of_readiness_to_cut': percentage_of_readiness_to_cut,
                                'number_of_details_plan': number_of_details_plan,
                                'number_of_details_fact': number_of_details_fact,
                                }
                                }
                            }
                if orders_data.get(composite_key) \
                        and orders_data[composite_key]['child_orders']:
                    orders_data[composite_key]['child_orders']['job monitor for work centers'][key] = {
                            'percentage_of_readiness_to_cut': percentage_of_readiness_to_cut,
                            'number_of_details_plan': number_of_details_plan,
                            'number_of_details_fact': number_of_details_fact,
                    }
    return orders_data


def extract_data_to_report_production_orders_report(orders_data: dict,
                                                    workbook: object,
                                                    sheet: str) -> dict:
    sub_orders_temp = dict()
    workbook_sheet = workbook[sheet]

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True),
                      ncols=80, ascii=True, desc=sheet):
        full_order_date = value[0]
        order_date = full_order_date[6:10]
        order_number = "{:05d}".format(value[1])
        composite_key = order_date + order_number
        order_company = value[2]
        order_division = value[3]
        order_launch_date = value[4]
        order_execution_date = value[5]
        responsible = value[6]
        comment = value[7]
        order_parent = value[8]

        # Если в строке есть основной заказ, то это дополнительный заказ на раскрой и покраску
        # Собираем информацию о дополнительных заказах
        if order_parent:
            composite_key_parent = order_parent[43:47] + order_parent[28:33]
            if orders_data.get(composite_key_parent) \
                    and not orders_data[composite_key_parent].report_sub_order:
                sub_orders_temp[composite_key] = composite_key_parent
                report_sub_order = ReportSubOrder()
                report_description = ReportDescriptionOfProductionOrder()
                report_description.composite_key = composite_key
                report_description.order_date = full_order_date[:10]
                report_description.order_number = order_number
                report_description.order_company = order_company
                report_description.order_division = order_division
                report_description.order_launch_date = order_launch_date
                report_description.order_execution_date = order_execution_date
                report_description.responsible = responsible
                report_description.comment = comment
                orders_data[composite_key_parent].report_sub_order = report_sub_order
                orders_data[composite_key_parent].report_sub_order.report_description_of_production_orders[composite_key] = report_description

            if orders_data.get(composite_key_parent) \
                    and orders_data[composite_key_parent].report_sub_order:
                sub_orders_temp[composite_key] = composite_key_parent
                report_description = ReportDescriptionOfProductionOrder()
                report_description.composite_key = composite_key
                report_description.order_date = full_order_date[:10]
                report_description.order_number = order_number
                report_description.order_company = order_company
                report_description.order_division = order_division
                report_description.order_launch_date = order_launch_date
                report_description.order_execution_date = order_execution_date
                report_description.responsible = responsible
                report_description.comment = comment
                orders_data[composite_key_parent].report_sub_order.report_description_of_production_orders[composite_key] = report_description

            if sub_orders_temp.get(composite_key_parent) \
                and orders_data.get(sub_orders_temp[composite_key_parent]) \
                    and orders_data[sub_orders_temp[composite_key_parent]].report_sub_order:
                report_description = ReportDescriptionOfProductionOrder()
                report_description.composite_key = composite_key
                report_description.order_date = full_order_date[:10]
                report_description.order_number = order_number
                report_description.order_company = order_company
                report_description.order_division = order_division
                report_description.order_launch_date = order_launch_date
                report_description.order_execution_date = order_execution_date
                report_description.responsible = responsible
                report_description.comment = comment
                orders_data[sub_orders_temp[composite_key_parent]].report_sub_order.report_description_of_production_orders[composite_key] = report_description

            if sub_orders_temp.get(composite_key_parent) \
                and orders_data.get(sub_orders_temp[composite_key_parent]) \
                    and not orders_data[sub_orders_temp[composite_key_parent]].report_sub_order:
                sub_orders_temp[composite_key] = composite_key_parent
                report_description = ReportDescriptionOfProductionOrder()
                report_description.composite_key = composite_key
                report_description.order_date = full_order_date[:10]
                report_description.order_number = order_number
                report_description.order_company = order_company
                report_description.order_division = order_division
                report_description.order_launch_date = order_launch_date
                report_description.order_execution_date = order_execution_date
                report_description.responsible = responsible
                report_description.comment = comment
                report_sub_order = ReportSubOrder()
                orders_data[sub_orders_temp[composite_key_parent]].report_sub_order = report_sub_order
                orders_data[sub_orders_temp[composite_key_parent]].report_sub_order.report_description_of_production_orders[composite_key] = report_description

        # Если строка пуста, то это основной заказ
        # Собираем данные по описанию основного заказа на производство
        else:
            if orders_data.get(composite_key):
                report_description = ReportDescriptionOfProductionOrder()
                report_description.composite_key = composite_key
                report_description.order_date = full_order_date[:10]
                report_description.order_number = order_number
                report_description.order_company = order_company
                report_description.order_division = order_division
                report_description.order_launch_date = order_launch_date
                report_description.order_execution_date = order_execution_date
                report_description.responsible = responsible
                report_description.comment = comment
                orders_data[composite_key].report_description_main_order = report_description
    return orders_data


def extract_data_production_orders_report(orders_data: dict,
                                          workbook: object,
                                          sheet: str,
                                          config: object) -> dict:
    """Функция предназначена для сбора данных из отчета "Заказы на производство"

    Args:
        orders_data (dict): [данные по заказам на производство]
        workbook (object): [файл excel с исходными данными]
        sheet (str): [наименование листа]

    Returns:
        orders_data (dict): [данные по заказам на производство]
    """
    child_orders_temp = dict()
    workbook_sheet = workbook[sheet]

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True),
                      ncols=80, ascii=True, desc=config.sheet_division_1C):
        full_order_date = value[0]
        order_date = full_order_date[6:10]
        order_number = "{:05d}".format(value[1])
        composite_key = order_date + order_number
        order_company = value[2]
        order_division = value[3]
        order_launch_date = value[4]
        order_execution_date = value[5]
        responsible = value[6]
        comment = value[7]
        order_parent = value[8]

        # Если в строке есть основной заказ, то это дополнительный заказ на раскрой и покраску
        # Собираем информацию о дополнительных заказах
        if order_parent:
            composite_key_parent = order_parent[43:47] + order_parent[28:33]
            if orders_data.get(composite_key_parent) \
                    and not orders_data[composite_key_parent].get('child_orders'):
                child_orders_temp[composite_key] = composite_key_parent
                orders_data[composite_key_parent]['child_orders'] = {
                    'report description of production orders': {
                        composite_key: {
                            'order_date': order_date,
                            'order_number': order_number,
                            'order_company': order_company,
                            'order_division': order_division,
                            'order_launch_date': order_launch_date,
                            'order_execution_date': order_execution_date,
                            'responsible': responsible,
                            'comment': comment,
                            }
                            }
                            }
            if orders_data.get(composite_key_parent) \
                    and orders_data[composite_key_parent]['child_orders'].get('report description of production orders'):
                child_orders_temp[composite_key] = composite_key_parent
                orders_data[composite_key_parent]['child_orders']['report description of production orders'][composite_key] = {
                    'order_date': order_date,
                    'order_number': order_number,
                    'order_company': order_company,
                    'order_division': order_division,
                    'order_launch_date': order_launch_date,
                    'order_execution_date': order_execution_date,
                    'responsible': responsible,
                    'comment': comment,
                    }
            if orders_data.get(composite_key_parent) \
                    and not orders_data[composite_key_parent]['child_orders'].get('report description of production orders'):
                child_orders_temp[composite_key] = composite_key_parent
                orders_data[composite_key_parent]['child_orders']['report description of production orders'] = {
                    composite_key: {
                        'order_date': order_date,
                        'order_number': order_number,
                        'order_company': order_company,
                        'order_division': order_division,
                        'order_launch_date': order_launch_date,
                        'order_execution_date': order_execution_date,
                        'responsible': responsible,
                        'comment': comment,
                        }
                }
            if child_orders_temp.get(composite_key_parent) \
                and orders_data.get(child_orders_temp[composite_key_parent]) \
                    and orders_data[child_orders_temp[composite_key_parent]]['child_orders'].get('report description of production orders'):
                orders_data[child_orders_temp[composite_key_parent]]['child_orders']['report description of production orders'][composite_key] = {
                    'order_date': order_date,
                    'order_number': order_number,
                    'order_company': order_company,
                    'order_division': order_division,
                    'order_launch_date': order_launch_date,
                    'order_execution_date': order_execution_date,
                    'responsible': responsible,
                    'comment': comment,
                    }
            if child_orders_temp.get(composite_key_parent) \
                and orders_data.get(child_orders_temp[composite_key_parent]) \
                    and not orders_data[child_orders_temp[composite_key_parent]]['child_orders'].get('report description of production orders'):
                child_orders_temp[composite_key] = composite_key_parent
                orders_data[child_orders_temp[composite_key_parent]]['child_orders']['report description of production orders'] = {
                    composite_key: {
                        'order_date': order_date,
                        'order_number': order_number,
                        'order_company': order_company,
                        'order_division': order_division,
                        'order_launch_date': order_launch_date,
                        'order_execution_date': order_execution_date,
                        'responsible': responsible,
                        'comment': comment,
                        }
                }
        # Если строка пуста, то это основной заказ
        # Собираем данные по описанию основного заказа на производство
        else:
            if orders_data.get(composite_key):
                orders_data[composite_key]['report description of production orders'] = {
                        'order_date': order_date,
                        'order_number': order_number,
                        'order_company': order_company,
                        'order_division': order_division,
                        'order_launch_date': order_launch_date,
                        'order_execution_date': order_execution_date,
                        'responsible': responsible,
                        'comment': comment,
                    }
    return orders_data


class Bunch(dict):
    def __init__(self, *args, ** kwds):
        super(Bunch, self).__init__(*args, **kwds)
        self.__dict__ = self


def calculation_percentage_of_assembly(ordered: int,
                           released: int,
                           assembly_shop: int) -> float:
    """Функция расчитывает процент готовности сборки

    Args:
        ordered (int): [Количество изделий в заказе]
        released (int): [Выпущено комплектов]
        assembly_shop (int): [Выпуск сборка, отчет мастеров]

    Returns:
        float: [Процент готовности заказа по сборке]
    """
    assembly = assembly_shop if assembly_shop > released else released
    percentage = assembly / ordered
    if percentage > 1:
        percentage = 1
    return percentage


def determine_if_there_is_a_painting(orders_report: object,
                                     product_is_painted: dict,
                                     session_db: object,
                                     table_in_db: object) -> dict:
    """Функция определяет является изделие крашенным или нет.
    Если изделие красится, то дополнительно определяется есть
    в заказе на раскрой не крашенные детали.

    Args:
        orders_report (object): [Итоговый отчет для файла]
        product_is_painted (dict): [Словарь для кэширования статуса]
        session_db (object): [Открытый сеанс с БД]
        table_in_db (object): [Таблица в базе данных с описанием заказов]

    Returns:
        dict: [Словарь со статусом для painted_status: крашеное/не крашеное
                                   для cutting_status: есть корпус/ нет корпуса]
    """
    for order in tqdm(orders_report, ncols=80, ascii=True, desc="Присвоить статус крешенным издлиям"):
        furniture_article = order.furniture_article
        sub_orders_descript = session_db.query(table_in_db).filter(table_in_db.order_id == order.id)
        if sub_orders_descript:
            painted = False
            not_painted = 0
            for sub_order_descript in sub_orders_descript:
                if sub_order_descript.order_division == "Цех покраски":
                    painted = True
                if sub_order_descript.order_division == "Цех раскроечный":
                    not_painted += 1
            if not_painted == 1 and not painted:
                painted_status = "н/к"
                cutting_status = ""
                product_is_painted[furniture_article] = {
                    'painted_status': painted_status,
                    'cutting_status': cutting_status,
                }
            if not_painted == 1 and painted:
                painted_status = "к"
                cutting_status = "нет корпуса"
                product_is_painted[furniture_article] = {
                    'painted_status': painted_status,
                    'cutting_status': cutting_status,
                }
            if not_painted == 2 and painted:
                painted_status = "к"
                cutting_status = "есть корпус"
                product_is_painted[furniture_article] = {
                    'painted_status': painted_status,
                    'cutting_status': cutting_status,
                }
    return product_is_painted


def calculate_percentage_of_painting_readiness(
        painted_status: str,
        cutting_shop_for_painting: int,
        paint_shop_for_assembly: int) -> float:
    """Функция расчета процента готовности покраски

    Args:
        painted_status (str): [Статус изделия: к (крашеное)]
        cutting_shop_for_painting (int): [Выпуск комплектов раскрой
                                          на покраску]
        paint_shop_for_assembly (int): [Выпуск комплектов покраска]

    Returns:
        float: [Процент готовности покраски]
    """
    percentage_of_painting = ""
    if painted_status == "к" and cutting_shop_for_painting:
        percentage_of_painting = paint_shop_for_assembly / cutting_shop_for_painting
    return percentage_of_painting


def calculation_number_details_fact_paint_to_assembly(
        percentage_of_readiness_painting: float,
        number_of_details_plan_cut_to_paint: int) -> int:
    """Расчет количества деталей покраски на буфер.

    Это не фактические данные учета, а расчетные.

    Расчет ведется из учета деталей переданных на участок покраски
    цехом раскрой и количеством переданных комплектов покраски на буфер.

    Args:
        percentage_of_readiness_painting (float): [Процент готовности покраски]
        number_of_details_plan_cut_to_paint (int): [Количество деталей (план)
                                                    переданных участком раскроя
                                                    на покраску]

    Returns:
        int: [Количество деталей выпущенных цехом покраски на буфер]
    """
    number_of_details = ""
    if percentage_of_readiness_painting:
        number_of_details = int(percentage_of_readiness_painting * number_of_details_plan_cut_to_paint)
    return number_of_details


def extract_data_contractors(contractors: dict,
                             workbook: object,
                             config: object) -> dict:
    """Функция извлекает данные из файла Excel
    какое изделие изготавливается для какого контрагента

    Args:
        contractors (dict): [Контрагенты]
        workbook (object): [Файл Excel с данными по контрагентам]
        config (object): [Класс с конфигами]

    Returns:
        dict: [key: артикул, item: контрагент]
    """
    for value in tqdm(workbook.iter_rows(min_row=2, max_col=2), ncols=80, ascii=True, desc="'Извлекаем данные по контагентам"):
        key = value[0].value
        contractor = value[1].value
        if not contractors.get(key):
            contractors[key] = contractor
    return contractors


def set_formula_to_cell(formula: str,
                        formula_param: list,
                        worksheet: object,
                        row: int,
                        start_row: int,
                        last_row: int,
                        columns_number: dict,
                        columns_with_formula: list,
                        converter_letter: object) -> None:
    """Функция вставляет в строке формулу в столбцах перечисленных
    в списке

    Args:
        formula (str): [Наименование формулы: прим.: SUBTOTAL]
        formula_param (list): [Параметры формулы: [9] или [9, 3, ... 5]]
        worksheet (object): [Рабочий лист файла excel]
        row (int): [Строка в которую необходимо втавить формулу]
        start_row (int): [Стартовая строка для вычислений формулы]
        last_row (int): [Последняя строка для вычислений формулы]
        columns_number (dict): [Номер колонки: key= "Имя", item= int]
        columns_with_formula (list): [Названия колонок куда необходимо
                                      вставить формулы]
        converter_letter (object): [Конвертер перевода номера в символ.
                              Пример: колонка с номером 2, convert = B]

    Returns:
        str: [Сообщение: Формула ФОРМУЛА вставлена в (кол-во) ячеек]
    """
    for coll_name in tqdm(columns_with_formula,
                          ncols=80, ascii=True,
                          desc=f"Вставляем формулу {formula} в {len(columns_with_formula)} ячеек."):
        column_letter = converter_letter(columns_number[coll_name])
        worksheet.cell(row=row, column=columns_number[coll_name]).value = \
            f"={formula}({formula_param[0]},{column_letter}{start_row}:{column_letter}{last_row})"


def set_format_to_cell(format_cell: str,
                       worksheet: object,
                       start_row: int,
                       last_row: int,
                       columns_number: dict,
                       columns_with_format: list,
                       converter_letter: object) -> None:
    for coll_name in tqdm(columns_with_format,
                          ncols=80, ascii=True,
                          desc=f"Задаем формат {format_cell} для стобцов"):
        column_letter = converter_letter(columns_number[coll_name])
        for cell in worksheet[f"{column_letter}{start_row}:{column_letter}{last_row}"]:
            cell[0].number_format = format_cell


def set_weigth_to_cell(worksheet,
                       columns_number: dict,
                       columns_with_format: dict,
                       converter_letter: object) -> None:
    for key in tqdm(columns_with_format,
                    ncols=80, ascii=True,
                    desc=f'Задаем ширину для {len(columns_with_format)} столбцов.'):
        column_letter = converter_letter(columns_number[key])
        worksheet.column_dimensions[f'{column_letter}'].width = columns_with_format[key]


def set_styles_to_cells(worksheet,
                        start_row: int,
                        last_row: int,
                        column_name_left: str,
                        column_name_right: str,
                        columns_number: dict,
                        converter_letter: object,
                        styles: str) -> None:
    column_letter_left = converter_letter(columns_number[column_name_left])
    column_letter_right = converter_letter(columns_number[column_name_right])
    for cells in tqdm(worksheet[f'{column_letter_left}{start_row}:{column_letter_right}{last_row}'],
                      ncols=80, ascii=True,
                      desc=f'Задаем стили {styles} для диапазона ячеек от {column_name_left} до {column_name_right}.'):
        for cell in cells:
            cell.style = styles


def determine_ready_status_of_assembly(released: int,
                                       cutting_shop_for_assembly: int,
                                       cutting_shop_for_painting: int,
                                       paint_shop_for_assembly: int,
                                       painted_status: str,
                                       cutting_status: str,
                                       percentg_of_assembly: float,
                                       percentage_of_readiness_to_cut: float) -> tuple:
    """Функция определяет статус готовности заказа и расчитывает количество
    комплектов готовых к сборке.
    Статусы:
            Готов
            Готов к сборке
            На сборке
            В работе
            Развернут
            Не развернут

    Args:
        released (int): [Выпущено комплектов на склад ГП]
        cutting_shop_for_assembly (int): [Выпущено комплектов с раскроя на буфер]
        cutting_shop_for_painting (int): [Выпущено комплектов с раскроя на покраску]
        paint_shop_for_assembly (int): [Выпущено комплектов с покраски]
        painted_status (str): [Статус покраски: крашеное/не крашеное]
        cutting_status (str): [Статус корпус: есть корпус/нет корпуса]
        percentg_of_assembly (float): [Процент готовности по сборке]
        percentage_of_readiness_to_cut (float): [Процент готовности по раскрою]

    Returns:
        tuple [str, int]: [assembly_ready_status = Статус готовности (str),
                           quantity_to_be_assembled = количество (int)]
    """
    assembly_ready_status = ""
    quantity_to_be_assembled = ""

    # Статус готов
    if percentg_of_assembly == 1:
        assembly_ready_status = "Готов"
        quantity_to_be_assembled = ""
        return assembly_ready_status, quantity_to_be_assembled

    # Не крашеное / процент готовности 0%
    if painted_status == "н/к" \
        and cutting_shop_for_assembly \
            and percentg_of_assembly == 0:
        assembly_ready_status = "Готов к сборке"
        quantity_to_be_assembled = cutting_shop_for_assembly
        return assembly_ready_status, quantity_to_be_assembled

    # Не крашеное / процент готовности > 0%
    if painted_status == "н/к" \
        and cutting_shop_for_assembly \
            and 0 < percentg_of_assembly < 0.999:
        assembly_ready_status = "На сборке"
        quantity_to_be_assembled = (cutting_shop_for_assembly - released) \
            if (cutting_shop_for_assembly - released) > 0 else ""
        return assembly_ready_status, quantity_to_be_assembled

    # Крашеное c корпусом/ процент готовности 0%
    if painted_status == "к" \
        and cutting_status == "есть корпус" \
            and cutting_shop_for_assembly \
                and paint_shop_for_assembly \
                    and percentg_of_assembly == 0:
        assembly_ready_status = "Готов к сборке"
        quantity_to_be_assembled = min(cutting_shop_for_assembly, paint_shop_for_assembly)
        return assembly_ready_status, quantity_to_be_assembled

    # Крашеное c корпусом/ процент готовности > 0%
    if painted_status == "к" \
        and cutting_status == "есть корпус" \
            and cutting_shop_for_assembly \
                and paint_shop_for_assembly \
                    and 0 < percentg_of_assembly < 0.999:
        assembly_ready_status = "На сборке"
        quantity_to_be_assembled = min(cutting_shop_for_assembly, paint_shop_for_assembly) - released \
            if (min(cutting_shop_for_assembly, paint_shop_for_assembly) - released) > 0 else ""
        return assembly_ready_status, quantity_to_be_assembled

    # Крашеное без корпуса / процент готовности 0%
    if painted_status == "к" \
        and cutting_status == "нет корпуса" \
            and paint_shop_for_assembly \
                and percentg_of_assembly == 0:
        assembly_ready_status = "Готов к сборке"
        quantity_to_be_assembled = min(cutting_shop_for_painting ,paint_shop_for_assembly)
        return assembly_ready_status, quantity_to_be_assembled

    # Крашеное без корпуса / процент готовности > 0%
    if painted_status == "к" \
        and cutting_status == "нет корпуса" \
            and paint_shop_for_assembly \
                and 0 < percentg_of_assembly < 0.999:
        assembly_ready_status = "На сборке"
        quantity_to_be_assembled = min(cutting_shop_for_painting ,paint_shop_for_assembly) - released \
            if (min(cutting_shop_for_painting ,paint_shop_for_assembly) - released) > 0 else ""
        return assembly_ready_status, quantity_to_be_assembled

    # Развёрнутые заказы
    if percentg_of_assembly == 0 \
            and percentage_of_readiness_to_cut == 0:
        assembly_ready_status = "Развернут"
        return assembly_ready_status, quantity_to_be_assembled
    
    # Не развурнутые заказы
    if percentg_of_assembly == 0 \
            and percentage_of_readiness_to_cut == "не развернут" \
                and not cutting_shop_for_assembly \
                    and not cutting_shop_for_painting \
                        and not paint_shop_for_assembly:
        assembly_ready_status = "Не развернут"
        return assembly_ready_status, quantity_to_be_assembled

    # В работе раскрой покраска
    if not type(percentage_of_readiness_to_cut) == str:
        if percentg_of_assembly == 0 \
            and percentage_of_readiness_to_cut > 0 \
                and not assembly_ready_status == "Готов" \
                    and not assembly_ready_status == "Готов к сборке" \
                        and not assembly_ready_status == "На сборке" \
                            and not assembly_ready_status == "Развернут" \
                                and not assembly_ready_status == "Не развернут":
            assembly_ready_status = "В работе"
            return assembly_ready_status, quantity_to_be_assembled

    return assembly_ready_status, quantity_to_be_assembled


def calc_number_products_cutting_and_painting_workshops(
        ordered: int,
        released: int,
        remains_to_release: int,
        cutting_shop_for_assembly: int,
        cutting_shop_for_painting: int,
        paint_shop_for_assembly: int,
        painted_status: str,
        cutting_status: str,
        percentg_of_assembly: float,
        percentage_of_readiness_to_cut: float,
        assembly_ready_status: str) -> tuple[int, int, int]:
    cut_to_the_buffer_in_progress = 0
    cutting_for_painting_in_progress = 0
    painting_in_progress = 0

    # Не крашеные в работе
    if painted_status == "н/к" and (assembly_ready_status == "В работе" or assembly_ready_status == "Готов к сборке"):
        cut_to_the_buffer_in_progress = ordered - cutting_shop_for_assembly
        return cut_to_the_buffer_in_progress, \
            cutting_for_painting_in_progress, painting_in_progress

    # Крашеные c корпусом в работе
    if painted_status == "к" and cutting_status == "есть корпус" \
        and (assembly_ready_status == "В работе" or assembly_ready_status == "Готов к сборке"):
            if cutting_shop_for_assembly < ordered:
                cut_to_the_buffer_in_progress = ordered - cutting_shop_for_assembly
            if cutting_shop_for_painting < ordered:
                cutting_for_painting_in_progress = ordered - cutting_shop_for_painting
            if paint_shop_for_assembly < cutting_shop_for_painting:
                painting_in_progress = ordered - paint_shop_for_assembly
            return cut_to_the_buffer_in_progress, \
        cutting_for_painting_in_progress, painting_in_progress

    # Крашеные без корпуса в работе
    if painted_status == "к" and cutting_status == "нет корпуса" \
        and (assembly_ready_status == "В работе" or assembly_ready_status == "Готов к сборке"):
            if cutting_shop_for_painting < ordered:
                cutting_for_painting_in_progress = ordered - cutting_shop_for_painting
            if paint_shop_for_assembly < cutting_shop_for_painting:
                painting_in_progress = ordered - paint_shop_for_assembly
            return cut_to_the_buffer_in_progress, \
        cutting_for_painting_in_progress, painting_in_progress
    
    # Не крашеные на сборке
    if painted_status == "н/к" and assembly_ready_status == "На сборке":
        if cutting_shop_for_assembly - released < remains_to_release \
            and released < cutting_shop_for_assembly:
            cut_to_the_buffer_in_progress = remains_to_release - (cutting_shop_for_assembly - released)
            return cut_to_the_buffer_in_progress, \
                cutting_for_painting_in_progress, painting_in_progress
    
    # Крашеные с корпусом на сборке
    if painted_status == "к" and cutting_status == "есть корпус" \
        and assembly_ready_status == "На сборке":
            if cutting_shop_for_assembly - released < remains_to_release \
                and released < cutting_shop_for_assembly:
                    cut_to_the_buffer_in_progress = remains_to_release - (cutting_shop_for_assembly - released)
            if cutting_shop_for_painting - released < remains_to_release \
                and released < cutting_shop_for_painting:
                cutting_for_painting_in_progress = remains_to_release - (cutting_shop_for_painting - released)
            if paint_shop_for_assembly - released < remains_to_release \
                and released < paint_shop_for_assembly:
                painting_in_progress = remains_to_release - (paint_shop_for_assembly - released)
            return cut_to_the_buffer_in_progress, \
                cutting_for_painting_in_progress, painting_in_progress
    
    # Крешеные без корпуса на сборке
    if painted_status == "к" and cutting_status == "нет корпуса" \
        and assembly_ready_status == "На сборке":
            if cutting_shop_for_painting - released < remains_to_release \
                and released < cutting_shop_for_painting:
                cutting_for_painting_in_progress = remains_to_release - (cutting_shop_for_painting - released)
            if paint_shop_for_assembly - released < remains_to_release \
                and released < paint_shop_for_assembly:
                painting_in_progress = remains_to_release - (paint_shop_for_assembly - released)
            return cut_to_the_buffer_in_progress, \
                cutting_for_painting_in_progress, painting_in_progress

    return cut_to_the_buffer_in_progress, \
        cutting_for_painting_in_progress, painting_in_progress

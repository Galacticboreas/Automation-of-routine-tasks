import re

from tqdm import tqdm

from app.db_models.orders import (Report, ReportDescriptionOfProductionOrder,
                                  ReportMainOrder, ReportMonitorForWorkCenter,
                                  ReportMovingSetsOfFurniture,
                                  ReportReleaseOfAssemblyKits, ReportSubOrder)


def extract_data_to_report_moving_sets_of_furnuture(orders_data: dict,
                                                    workbook: object,
                                                    sheet: str,
                                                    expression: str,) -> dict:
    workbook_sheet = workbook[sheet]

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True), ncols=80, ascii=True, desc=sheet):
        full_order_number = value[0]
        composite_key = full_order_number[43:47] + full_order_number[28:33]
        furniture_name = value[1]
        ordered = value[2]
        released = value[3]
        remains_to_release = value[4]
        if (expression in full_order_number) and (ordered is not None):
            # Заполняем данные основного заказа
            order_main = ReportMainOrder()
            order_main.full_order_number = full_order_number
            order_main.furniture_name = furniture_name
            order_main.furniture_article = Report.extractor.get_article(furniture_name)
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

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True), ncols=80, ascii=True, desc=config.sheet_moving_1C):
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
                    'troubleshooting steps': 'отправить номер заказа в отдел учета для корректировки перемещения',
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

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True), ncols=80, ascii=True, desc=sheet):
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

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True), ncols=80, ascii=True, desc=config.sheet_kits_1C):
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

    for value in tqdm(workbook_sheet.iter_rows(min_row=2), ncols=80, ascii=True, desc=sheet):
        pattern = r'[0-9]{11}[" "]["о"]["т"][" "][0-9]{2}["."][0-9]{2}["."][0-9]{4}[" "][0-9]{2}[":"][0-9]{2}[":"][0-9]{2}'
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
                orders_data[composite_key].report_monitor_for_work_center = report_monitor
        else:
            sub_order = re.search(pattern, value[0].value)
            if sub_order:
                order_number = sub_order[0]
                key = order_number[21:25] + order_number[6:11]
                percentage_of_readiness_to_cut = value[2].value
                number_of_details_plan = value[3].value
                number_of_details_fact = value[4].value

                if orders_data.get(composite_key) and not orders_data[composite_key].report_sub_order:
                    # Заполняем процент готовности подзаказа
                    report_monitor = ReportMonitorForWorkCenter()
                    report_monitor.percentage_of_readiness_to_cut = percentage_of_readiness_to_cut
                    report_monitor.number_of_details_plan = number_of_details_plan
                    report_monitor.number_of_details_fact = number_of_details_fact
                    # Заполняем отчет для подзаказа по ключу
                    report_sub_order = ReportSubOrder()
                    orders_data[composite_key].report_sub_order = report_sub_order
                    orders_data[composite_key].report_sub_order.report_monitor_for_work_center[key] = report_monitor

                if orders_data.get(composite_key) and orders_data[composite_key].report_sub_order:
                    # Заполняем процент готовности подзаказа
                    report_monitor = ReportMonitorForWorkCenter()
                    report_monitor.percentage_of_readiness_to_cut = percentage_of_readiness_to_cut
                    report_monitor.number_of_details_plan = number_of_details_plan
                    report_monitor.number_of_details_fact = number_of_details_fact
                    # Заполняем отчет для подзаказа по ключу
                    orders_data[composite_key].report_sub_order.report_monitor_for_work_center[key] = report_monitor
    return orders_data


def extract_data_job_monitor_for_work_centers(orders_data: dict, workbook: object, sheet: str, config: object) -> dict:
    """Функция предназначена для сбора данных из отчета "Монитор рабочих центров"

    Args:
        orders_data (dict): [данные по заказам на производство]
        workbook (object): [файл excel с исходными данными]
        sheet (str): [наименование листа]

    Returns:
        orders_data (dict): [данные по заказам на производство]
    """
    workbook_sheet = workbook[sheet]

    for value in tqdm(workbook_sheet.iter_rows(min_row=2), ncols=80, ascii=True, desc=config.sheet_percentage_1C):
        pattern = r'[0-9]{11}[" "]["о"]["т"][" "][0-9]{2}["."][0-9]{2}["."][0-9]{4}[" "][0-9]{2}[":"][0-9]{2}[":"][0-9]{2}'
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

                if orders_data.get(composite_key) and not orders_data[composite_key].get('child_orders'):
                    orders_data[composite_key]['child_orders'] = {
                        'job monitor for work centers': {
                            key: {
                                'percentage_of_readiness_to_cut': percentage_of_readiness_to_cut,
                                'number_of_details_plan': number_of_details_plan,
                                'number_of_details_fact': number_of_details_fact,
                                }
                                }
                            }
                if orders_data.get(composite_key) and orders_data[composite_key]['child_orders']:
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

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True), ncols=80, ascii=True, desc=sheet):
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
            if orders_data.get(composite_key_parent) and not orders_data[composite_key_parent].report_sub_order:
                sub_orders_temp[composite_key] = composite_key_parent
                report_sub_order = ReportSubOrder()
                report_description = ReportDescriptionOfProductionOrder()
                report_description.composite_key = composite_key
                report_description.order_date = order_date
                report_description.order_number = order_number
                report_description.order_company = order_company
                report_description.order_division = order_division
                report_description.order_launch_date = order_launch_date
                report_description.order_execution_date = order_execution_date
                report_description.responsible = responsible
                report_description.comment = comment
                orders_data[composite_key_parent].report_sub_order = report_sub_order
                orders_data[composite_key_parent].report_sub_order.report_description_of_production_orders[composite_key] = report_description

            if orders_data.get(composite_key_parent) and orders_data[composite_key_parent].report_sub_order:
                sub_orders_temp[composite_key] = composite_key_parent
                report_description = ReportDescriptionOfProductionOrder()
                report_description.composite_key = composite_key
                report_description.order_date = order_date
                report_description.order_number = order_number
                report_description.order_company = order_company
                report_description.order_division = order_division
                report_description.order_launch_date = order_launch_date
                report_description.order_execution_date = order_execution_date
                report_description.responsible = responsible
                report_description.comment = comment
                orders_data[composite_key_parent].report_sub_order.report_description_of_production_orders[composite_key] = report_description

            if sub_orders_temp.get(composite_key_parent) and orders_data.get(sub_orders_temp[composite_key_parent]) and orders_data[sub_orders_temp[composite_key_parent]].report_sub_order:
                report_description = ReportDescriptionOfProductionOrder()
                report_description.composite_key = composite_key
                report_description.order_date = order_date
                report_description.order_number = order_number
                report_description.order_company = order_company
                report_description.order_division = order_division
                report_description.order_launch_date = order_launch_date
                report_description.order_execution_date = order_execution_date
                report_description.responsible = responsible
                report_description.comment = comment
                orders_data[sub_orders_temp[composite_key_parent]].report_sub_order.report_description_of_production_orders[composite_key] = report_description

            if sub_orders_temp.get(composite_key_parent) and orders_data.get(sub_orders_temp[composite_key_parent]) and not orders_data[sub_orders_temp[composite_key_parent]].report_sub_order:
                sub_orders_temp[composite_key] = composite_key_parent
                report_description = ReportDescriptionOfProductionOrder()
                report_description.composite_key = composite_key
                report_description.order_date = order_date
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
                report_description.order_date = order_date
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

    for value in tqdm(workbook_sheet.iter_rows(min_row=2, values_only=True), ncols=80, ascii=True, desc=config.sheet_division_1C):
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
            if orders_data.get(composite_key_parent) and not orders_data[composite_key_parent].get('child_orders'):
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
            if orders_data.get(composite_key_parent) and orders_data[composite_key_parent]['child_orders'].get('report description of production orders'):
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
            if orders_data.get(composite_key_parent) and not orders_data[composite_key_parent]['child_orders'].get('report description of production orders'):
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
            if child_orders_temp.get(composite_key_parent) and orders_data.get(child_orders_temp[composite_key_parent]) and orders_data[child_orders_temp[composite_key_parent]]['child_orders'].get('report description of production orders'):
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
            if child_orders_temp.get(composite_key_parent) and orders_data.get(child_orders_temp[composite_key_parent]) and not orders_data[child_orders_temp[composite_key_parent]]['child_orders'].get('report description of production orders'):
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

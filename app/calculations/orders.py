import re

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

    for value in workbook_sheet.iter_rows(min_row=2, values_only=True):
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
                    'type of error' : 'ошибка учета',
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

def extract_data_release_of_assembly_kits(orders_data: dict,
                                          workbook: object,
                                          sheet: str,
                                          expression: str) -> dict:
    
    workbook_sheet = workbook[sheet]

    for value in workbook_sheet.iter_rows(min_row=2, values_only=True):
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

def extract_data_job_monitor_for_work_centers(orders_data: dict,
                                              workbook: object,
                                              sheet: str) -> dict:
    workbook_sheet = workbook[sheet]

    for value in workbook_sheet.iter_rows(min_row=2):
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

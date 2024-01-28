def extract_data_moving_sets_of_furniture(orders_data: dict,
                                          error_log: dict,
                                          workbook: object,
                                          sheet: str,
                                          expression: str,
                                          extractor: object,
                                          config: object) -> dict:
    sheet = workbook[config.sheet_moving_1C]

    for value in sheet.iter_rows(min_row=2, values_only=True):
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
                'moving_sets_of_furniture': {
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

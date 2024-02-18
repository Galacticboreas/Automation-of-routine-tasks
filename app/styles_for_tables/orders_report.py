from openpyxl.styles import (Alignment, Border, Font, NamedStyle, PatternFill,
                             Side)

# Стили для шапки таблицы

# Стиль шрифта
font = Font(
    name='Calibri',
    size=8,
    bold=False,
    italic=False,
    vertAlign=None,
    underline='none',
    strike=False,
    color='00000000',
)
# Заливка ячеек
fill = PatternFill(fill_type='solid', fgColor="0099CCFF")
# Границы ячеек
border = Border(
    left=Side(border_style='thin', color='FF000000'),
    right=Side(border_style='thin', color='FF000000'),
    top=Side(border_style='thin', color='FF000000'),
    bottom=Side(border_style='thin', color='FF000000'),
)
# Выравнивание в ячейках
alignment = Alignment(
    horizontal='center',
    vertical='center',
    text_rotation=0,
    wrap_text=True,
    shrink_to_fit=False,
    indent=0,
)
# Именованный стиль для шапки таблицы
t_head = NamedStyle(name='table_header')
t_head.font = font
t_head.fill = fill
t_head.border = border
t_head.alignment = alignment

# Стили для данных таблицы

# Стиль шрифта для текстовых данных
font = Font(
    name='Calibri',
    size=11,
    bold=False,
    italic=False,
    vertAlign=None,
    underline='none',
    strike=False,
    color='00000000',
)
# Границы ячеек
border = Border(
    left=Side(border_style='thin', color='FF000000'),
    right=Side(border_style='thin', color='FF000000'),
    top=Side(border_style='thin', color='FF000000'),
    bottom=Side(border_style='thin', color='FF000000'),
)
# Выравнивание в ячейках
alignment = Alignment(
    horizontal='left',
    vertical='center',
    text_rotation=0,
    wrap_text=False,
    shrink_to_fit=False,
    indent=0,
)
# Именованный стиль для шапки таблицы
td_text = NamedStyle(name='table_data_text')
td_text.font = font
td_text.border = border
td_text.alignment = alignment

# Стиль шрифта для цифровых данных
font = Font(
    name='Calibri',
    size=11,
    bold=False,
    italic=False,
    vertAlign=None,
    underline='none',
    strike=False,
    color='00000000',
)
# Границы ячеек
border = Border(
    left=Side(border_style='thin', color='FF000000'),
    right=Side(border_style='thin', color='FF000000'),
    top=Side(border_style='thin', color='FF000000'),
    bottom=Side(border_style='thin', color='FF000000'),
)
# Выравнивание в ячейках
alignment = Alignment(
    horizontal='right',
    vertical='center',
    text_rotation=0,
    wrap_text=False,
    shrink_to_fit=False,
    indent=0,
)
# Именованный стиль для шапки таблицы
td_digit = NamedStyle(name='table_data_digit')
td_digit.font = font
td_digit.border = border
td_digit.alignment = alignment

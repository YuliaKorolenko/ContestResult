from cfscript import ContestInfo
from gspread import Client, Spreadsheet, Worksheet, service_account
from gspread_formatting import GridRange, GradientRule, InterpolationPoint, Color, ConditionalFormatRule, get_conditional_format_rules
from typing import List, Dict

def client_init_json() -> Client:
    """Создание клиента для работы с Google Sheets."""
    return service_account(filename='bcdist.json')

def get_table_by_id(client: Client, table_url):
    """Получение таблицы из Google Sheets по ID таблицы."""
    return client.open_by_key(table_url)


def extract_data_from_sheet(table: Spreadsheet, sheet_name: str) -> List[List[str]]:
    """
    Извлекает данные из указанного листа таблицы Google Sheets и возвращает список словарей.

    :param table: Объект таблицы Google Sheets (Spreadsheet).
    :param sheet_name: Название листа в таблице.
    :return: Список словарей, представляющих данные из таблицы.
    """
    worksheet = table.worksheet(sheet_name)
    rows = worksheet.get_all_records()
    
    rows_to_insert = []
    for child in rows:
        row = [child.get("name", ""), child.get("cf_login", ""),]
        rows_to_insert.append(row)
    return rows_to_insert

def create_worksheet(table: Spreadsheet, title: str, rows: int, cols: int):
    """Создание листа в таблице."""
    return table.add_worksheet(title, rows, cols)

def set_column_width(worksheet, column_start, column_end, width):
    """
    Устанавливает ширину колонки в Google Sheets.

    :param worksheet: Объект Worksheet из gspread.
    :param column_index: Индекс колонки (начиная с 1).
    :param width: Ширина колонки в пикселях.
    """
    body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": column_start,  
                        "endIndex": column_end,      # не включительно
                    },
                    "properties": {
                        "pixelSize": width
                    },
                    "fields": "pixelSize"
                }
            }
        ]
    }
    # Отправка запроса
    worksheet.spreadsheet.batch_update(body)

def add_borders_to_range(worksheet, start_row, end_row, start_column, end_column):
    """Добавление двойных границ в указанный диапазон."""
    request = {
        "requests": [{
            "updateBorders": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": start_column,
                    "endColumnIndex": end_column,
                },
                "top": {"style": "DOUBLE"},
                "bottom": {"style": "DOUBLE"},
                "left": {"style": "DOUBLE"},
                "right": {"style": "DOUBLE"},
            }
        }]
    }
    worksheet.spreadsheet.batch_update(request)

def column_number_to_letter(col_num : int) -> str:
    """Преобразует номер столбца (1-based) в буквы Excel (A, B, ..., Z, AA, AB, ...)."""
    letters = ""
    while col_num > 0:
        col_num -= 1
        letters = chr(65 + (col_num % 26)) + letters
        col_num //= 26
    return letters


def insert_contest_info(worksheet: Worksheet, contest_info: ContestInfo, start_col: int = 1):
    """    
    :param worksheet: Объект Worksheet из gspread.
    :param contest_info: Объект ContestInfo с данными о контесте.
    :param start_col: Номер стартовой колонки для вставки данных (по умолчанию — 1).
    """
    len_task = len(contest_info.Tasks)
    end_col = start_col + len_task
    start_letter = column_number_to_letter(start_col)
    end_letter = column_number_to_letter(end_col)
    result_start_row = 3  # Результаты начинаются с третьей строки

    # Вставляем название контеста в первую строку (объединяем колонки)
    worksheet.update(f"{start_letter}1", [[contest_info.name]])
    insert_sum_formula(worksheet=worksheet, target_col=start_col, end_col=end_col, start_row=result_start_row, end_row=16)
    apply_gradient_formatting(worksheet=worksheet, column_letter=start_letter, start_row=result_start_row, end_row=16, mid_value=1 * len_task//3)
    cur_col = start_col + 1
    cur_letter = column_number_to_letter(cur_col)

    # Вставляем заголовки задач во вторую строку
    worksheet.update(f"{cur_letter}2:{end_letter}2", [contest_info.Tasks])
     
    # Вставляем результаты в последующие строки
    adresses = f"{cur_letter}{result_start_row}:{end_letter}{result_start_row + len(contest_info.Result) - 1}"
    worksheet.update(
        adresses,
        contest_info.Result
    )
    add_borders_to_range(worksheet, start_row=0, end_row=16, start_column=start_col - 1, end_column=end_col)
    set_column_width(worksheet=worksheet, column_start = start_col - 1, column_end = end_col, width=40)

def update_children_info(worksheet: Worksheet, children):
    end_col = 3
    children_count = len(children)

    header = [["ФИО", "Cf", "Σ"]]
    worksheet.update(f"A1:B1", [["Группа", "Группа"]]) 
    worksheet.update(f"A2:C2", header)  
    worksheet.update(f"A3:B{children_count + 2}", children) 

    add_borders_to_range(worksheet, start_row=0, end_row=1, start_column=0, end_column=end_col)
    add_borders_to_range(worksheet, start_row=1, end_row=children_count + 2, start_column=0, end_column=end_col)
    set_column_width(worksheet, 0, 1, 220)
    set_column_width(worksheet, 1, 2, 150)

def insert_sum_formula(worksheet, target_col, end_col, start_row, end_row):
    # Преобразуем номера столбцов в буквы
    target_col_letter = column_number_to_letter(target_col)
    sum_start_letter = column_number_to_letter(target_col + 1)
    sum_end_letter = column_number_to_letter(end_col)

    header_cell = f"{target_col_letter}2"
    worksheet.update_acell(header_cell, "Σ")
    # Формируем обновления для каждой строки
    for row in range(start_row, end_row + 1):
        target_cell = f"{target_col_letter}{row}"
        sum_range = f"{sum_start_letter}{row}:{sum_end_letter}{row}"
        formula = '=SUM(' + sum_range + ')'
        worksheet.update_acell(target_cell, formula)

    # apply_gradient_formatting(worksheet=worksheet, 
                            #   column_letter=target_col_letter, 
                            #   start_row=column_number_to_letter(start_row),
                            #   end_row=column_number_to_letter(end_row),
                            #   mid_value=15)


def apply_gradient_formatting(worksheet, column_letter, start_row, end_row, mid_value):
    """
    Применяет условное форматирование с градиентом для указанной колонки.

    :param worksheet: Объект Worksheet из gspread.
    :param column_letter: Буква колонки (например, 'A').
    :param start_row: Номер начальной строки.
    :param end_row: Номер конечной строки.
    """
    # Диапазон форматирования
    range_address = f"{column_letter}{start_row}:{column_letter}{end_row}"

    max_color = Color(87 / 256, 187 / 256, 138 / 256)  # Зеленый цвет
    min_color = Color(230 / 256, 124 / 256, 115 / 256) # Красный цвет
    mid_color = Color(255 / 256, 214 / 256, 102 / 256)

    # Правило цветового градиента
    rule = ConditionalFormatRule(
        ranges=[GridRange.from_a1_range(range_address, worksheet)],
        gradientRule=GradientRule(
            maxpoint=InterpolationPoint(color=max_color, type='MAX'),
            minpoint=InterpolationPoint(color=min_color, type='MIN'),
            midpoint=InterpolationPoint(color=mid_color, type="NUMBER", value=f"{mid_value}")
        )
    )
    rules = get_conditional_format_rules(worksheet)
    rules.append(rule)
    rules.save()


if __name__ == "__main__":
    client = client_init_json()
    table = get_table_by_id(client, table_id)




    


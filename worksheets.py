from cfscript import ContestInfo
from gspread import Client, Spreadsheet, Worksheet, service_account
from gspread_formatting import GridRange, GradientRule, InterpolationPoint, Color, ConditionalFormatRule, get_conditional_format_rules
from typing import List, Dict

import config

def client_init_json() -> Client:
    """Создание клиента для работы с Google Sheets."""
    return service_account(filename=config.GOOGLE_SA_PATH)

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
        row = [child.get("surname", ""), child.get("name", ""), child.get("cf_login", ""),]
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

def add_borders_to_range(worksheet, start_row, end_row, start_column, end_column, add_inner = False):
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
    if (add_inner):
        request["requests"][0]["updateBorders"].update({
            "innerHorizontal": {"style": "DOUBLE"},  
            "innerVertical": {"style": "DOUBLE"},   
        })
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
    END_ROW = len(contest_info.Result) + result_start_row - 1

    # Вставляем название контеста в первую строку (объединяем колонки)
    worksheet.update(f"{start_letter}1", [[contest_info.name]])
    insert_sum_formula(worksheet=worksheet, target_col=start_col, end_col=end_col, start_row=result_start_row, end_row=END_ROW)
    apply_gradient_formatting(worksheet=worksheet, column_letter=start_letter, start_row=result_start_row, end_row=END_ROW, mid_value=1 * len_task//3)
    cur_col = start_col + 1
    cur_letter = column_number_to_letter(cur_col)

    # Вставляем заголовки задач во вторую строку
    worksheet.update(f"{cur_letter}2:{end_letter}2", [contest_info.Tasks])
     
    # Вставляем результаты в последующие строки
    filtered_result = [["" if cell == -1 else cell for cell in row] for row in contest_info.Result]
    adresses = f"{cur_letter}{result_start_row}:{end_letter}{result_start_row + len(contest_info.Result) - 1}"
    worksheet.update(
        adresses,
        filtered_result
    )

    titleAdress = f"{start_letter}{1}:{end_letter}{1}"
    worksheet.merge_cells(titleAdress)  
    style_cells(worksheet=worksheet, cell_range=titleAdress, font_size=10)
    add_borders_to_range(worksheet, start_row=0, end_row=1, start_column=start_col - 1, end_column=end_col)
    style_cells(worksheet=worksheet, cell_range=f"{start_letter}{2}:{end_letter}{2}", font_size=10)
    add_borders_to_range(worksheet, start_row=1, end_row=2, start_column=start_col - 1, end_column=end_col, add_inner=True)

    add_borders_to_range(worksheet, start_row=0, end_row=END_ROW, start_column=start_col - 1, end_column=end_col)
    apply_gradient(worksheet=worksheet, column_letter=cur_letter, end_column=end_letter, start_row=result_start_row, end_row=END_ROW)
    set_column_width(worksheet=worksheet, column_start = start_col - 1, column_end = end_col, width=40)

def style_cells(worksheet: Worksheet, cell_range: str, font_size: int = 12):
    """
    :param worksheet: Worksheet объект
    :param cell_range: Диапазон ячеек в формате "A1:C3"
    :param font_size: Размер шрифта (по умолчанию 12)
    """
    worksheet.format(cell_range, {
        "horizontalAlignment": "CENTER",  # Центрирование по горизонтали
        "verticalAlignment": "MIDDLE",   # Центрирование по вертикали
        "textFormat": {
            "bold": True,          
            "fontSize": font_size,  
             "fontFamily": "Arial"  
        },
        "backgroundColor": {       # Серая заливка
            "red": 243 / 256,
            "green": 243 / 256,
            "blue": 243 / 256,
        }
    })

def update_children_info(worksheet: Worksheet, header, children):
    children_count = len(children)
    header_size = len(header[0])

    worksheet.update(f"A2:D2", header) 

    add_borders_to_range(worksheet=worksheet, start_row=0, end_row=2, start_column=0, end_column=header_size, add_inner=True)

    worksheet.merge_cells("A1:A2")  
    worksheet.merge_cells("B1:B2")  
    worksheet.merge_cells("C1:C2")  
    worksheet.merge_cells("D1:D2")  

    worksheet.update(f"A3:C{children_count + 2}", children) 
    style_cells(worksheet=worksheet, cell_range=f"{"A"}{1}:{"D"}{2}", font_size=13)

    add_borders_to_range(worksheet, start_row=0, end_row=1, start_column=0, end_column=header_size)
    add_borders_to_range(worksheet, start_row=1, end_row=children_count + 2, start_column=0, end_column=header_size)
    
    # Calculate max length for each column and set width accordingly
    # Include header row in calculation
    header_row = header[0]
    
    # Find max length in each column
    max_lengths = []
    for col_idx in range(header_size):
        max_len = 0
        
        for child in children:
            if col_idx < len(child):
                cell_value = str(child[col_idx])
                max_len = max(max_len, len(cell_value))
        
        max_lengths.append(max_len)
    

    for col_idx in range(header_size):
        calculated_width = max(100, min(400, max_lengths[col_idx] * 8))
        set_column_width(worksheet, col_idx, col_idx + 1, calculated_width)

    set_column_width(worksheet, 0, 1, 220)


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


def insert_main_sum_column(worksheet, sum_col, contest_sum_columns, start_row, end_row, num_contests):
    """
    Inserts SUM formula in the main Σ column (column D) that sums all contest Σ columns.
    Formula format: = E3 + I3 + M3 + ... where E, I, M are contest Σ columns.
    Also applies gradient conditional formatting to this column.
    
    :param worksheet: Worksheet object
    :param sum_col: Column number for the main Σ column (4 for column D)
    :param contest_sum_columns: List of column numbers that are Σ columns for each contest
    :param start_row: First row with data (3)
    :param end_row: Last row with data
    :param num_contests: Number of contests (for calculating midpoint)
    """
    sum_col_letter = column_number_to_letter(sum_col)
    
    # Build formula like = E3 + I3 + M3 + ...
    for row in range(start_row, end_row + 1):
        # Create list of cell references: E3, I3, M3, etc.
        cell_refs = [f"{column_number_to_letter(col)}{row}" for col in contest_sum_columns]
        # Join them with plus: E3 + I3 + M3
        sum_cells = " + ".join(cell_refs)
        formula = f'={sum_cells}'
        target_cell = f"{sum_col_letter}{row}"
        worksheet.update_acell(target_cell, formula)
    
    # Calculate midpoint for gradient (average tasks per contest * num contests / 3)
    # Estimate: assume ~5 tasks per contest on average
    estimated_tasks_per_contest = 5
    mid_value = estimated_tasks_per_contest * num_contests // 3
    
    # Apply center alignment to the Σ column
    sum_range = f"{sum_col_letter}{start_row}:{sum_col_letter}{end_row}"
    worksheet.format(sum_range, {
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    })
    
    apply_gradient_formatting(
        worksheet=worksheet,
        column_letter=sum_col_letter,
        start_row=start_row,
        end_row=end_row,
        mid_value=mid_value,
        is_rgb=False
    )


def apply_gradient(worksheet, column_letter, end_column, start_row, end_row):
    # Диапазон форматирования
    range_address = f"{column_letter}{start_row}:{end_column}{end_row}"

    max_color = Color(87 / 256, 187 / 256, 138 / 256)  # Зеленый цвет
    min_color = Color(255 / 256, 214 / 256, 102 / 256) # Оранжевый

    # Правило цветового градиента
    rule = ConditionalFormatRule(
        ranges=[GridRange.from_a1_range(range_address, worksheet)],
        gradientRule=GradientRule(
            maxpoint=InterpolationPoint(color=max_color, type='MAX'),
            minpoint=InterpolationPoint(color=min_color, type='MIN'),
        )
    )
    rules = get_conditional_format_rules(worksheet)
    rules.append(rule)
    rules.save()


def apply_gradient_formatting(worksheet, column_letter, start_row, end_row, mid_value, is_rgb = True):
    """
    Применяет условное форматирование с градиентом для указанной колонки.

    :param worksheet: Объект Worksheet из gspread.
    :param column_letter: Буква колонки (например, 'A').
    :param start_row: Номер начальной строки.
    :param end_row: Номер конечной строки.
    """
    # Диапазон форматирования
    range_address = f"{column_letter}{start_row}:{column_letter}{end_row}"
   
    max_color = Color(87 / 256, 187 / 256, 138 / 256)  # Green color
    min_color = Color(230 / 256, 124 / 256, 115 / 256) # Red color
    mid_color = Color(255 / 256, 214 / 256, 102 / 256)
    if not is_rgb:
        max_color = Color(87/256, 187/256, 138/256)   # Green (#57BB8A)
        min_color = Color(255/256, 255/256, 255/256)  # White
        mid_color = Color(171/256, 221/256, 196/256)  # Light green midpoint (#ABDDC4)

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




    


import time

import config
from cfscript import CodeforcesServer
from worksheets import insert_contest_info, client_init_json, get_table_by_id, create_worksheet, update_children_info,extract_data_from_sheet, insert_sum_formula, insert_main_sum_column

contests_ids = config.CONTEST_IDS
table_link = config.TABLE_LINK
table_id = config.TABLE_ID
worksheet_name = config.WORKSHEET_NAME
header = [["Фамилия", "Имя", "Cf", "Σ"]]

if __name__ == "__main__":
    if not contests_ids:
        raise RuntimeError("CONTEST_IDS is required in config.local.py (list of contest IDs).")
    if not table_id:
        raise RuntimeError("TABLE_ID is required in config.local.py (Google Sheet key).")
    client = client_init_json()
    cur_table = get_table_by_id(client, table_id)
    cfserver = CodeforcesServer()
    worksheet_names = [ws.title for ws in cur_table.worksheets()]
    if worksheet_name in worksheet_names:
        worksheet = cur_table.worksheet(worksheet_name)
        cur_table.del_worksheet(worksheet)

    create_worksheet(table=cur_table, title=worksheet_name, rows=20, cols=500)
    worksheet = cur_table.worksheet(worksheet_name)
    children = extract_data_from_sheet(table=cur_table, sheet_name="handles")
    map = {}
    i = 0
    for child in children:
        print(child)
        map[child[2]] = i
        i += 1

    # update_children_info(worksheet=worksheet, header=header, children=children)

    start = len(header[0]) + 1  # First contest starts at column E (5)
    # Track all contest Σ columns (each contest has its Σ column at start_col)
    contest_sum_columns = []  # List of column numbers that are Σ columns for each contest
    result_start_row = 3  # Results start at row 3
    num_children = len(children)
    end_row = result_start_row + num_children - 1
    
    for contest_id in contests_ids:
        time.sleep(20)
        contest = cfserver.generate_contest_info(contest_id=contest_id, niknames=map)
        insert_contest_info(worksheet = worksheet, contest_info = contest, start_col = start)
        print(len(contest.Result))
        # Track the Σ column for this contest (at start_col)
        contest_sum_columns.append(start)
        start += len(contest.Tasks) + 1  # Move to next contest (skip Σ column + tasks)
    
    time.sleep(20)
    # Insert main Σ column formula and apply gradient formatting
    if contest_sum_columns:
        sum_col = len(header[0])  # Column D (4)
        insert_main_sum_column(
            worksheet=worksheet,
            sum_col=sum_col,
            contest_sum_columns=contest_sum_columns,
            start_row=result_start_row,
            end_row=end_row,
            num_contests=len(contests_ids)
        )

    
    

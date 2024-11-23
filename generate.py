from cfscript import CodeforcesServer
from worksheets import insert_contest_info, client_init_json, get_table_by_id, create_worksheet, update_children_info,extract_data_from_sheet, insert_sum_formula
import time

contests_ids = [557376, 559375, 561279, 563232, 566926, 568555]

table_link = ...
table_id = ...
worksheet_name = "Контесты2"

if __name__ == "__main__":
    client = client_init_json()
    cur_table = get_table_by_id(client, table_id)
    cfserver = CodeforcesServer()

    create_worksheet(table=cur_table, title=worksheet_name, rows=20, cols=500)
    worksheet = cur_table.worksheet(worksheet_name)
    children = extract_data_from_sheet(table=cur_table, sheet_name="handles")
    map = {}
    i = 0
    for child in children:
        map[child[1]] = i
        i += 1

    update_children_info(worksheet=worksheet, children=children)

    start = 4
    for contest_id in contests_ids:
        time.sleep(20)
        contest = cfserver.generate_contest_info(contest_id=contest_id, niknames=map)
        insert_contest_info(worksheet = worksheet, contest_info = contest, start_col = start)
        print(len(contest.Tasks))
        start += len(contest.Tasks) + 1
        #  len(task) + sum result

    
    

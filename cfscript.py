import codeforces_api
from typing import List, Dict
from dataclasses import dataclass

import config

@dataclass
class ContestInfo:
    name: str
    Tasks: List[str]
    Result: List[List[int]]

class CodeforcesServer:
    def __init__(self): 
        self.cf = codeforces_api.CodeforcesApi(api_key=config.CF_API_KEY, secret=config.CF_API_SECRET)

    def generate_contest_info(self, contest_id : int, niknames : Dict[str, int]):
        first_key = next(iter(niknames))
        standings = self.cf.contest_standings(contest_id=contest_id, handles=[first_key])

        name = standings["contest"].name
        problems_ids = [problem.index for problem in standings["problems"]] 
        problems_dict = {problem: idx for idx, problem in enumerate(problems_ids)}
        print(problems_ids)

        status = self.cf.contest_status(contest_id)
        contest_result = [[-1 for _ in range(len(problems_ids))] for _ in range(len(niknames.keys()))]
        for row in status:
            if (row.author.members[0].handle in niknames):
                # print(row)
                # print(row.verdict)
                # print(row.author.members[0].handle)
                # print(row.problem.index)
                num_child = niknames[row.author.members[0].handle]
                problem_num = problems_dict[row.problem.index]
                if (row.verdict == "OK") :
                    contest_result[num_child][problem_num] = 1   
                elif (contest_result[num_child][problem_num] != 1) :
                    contest_result[num_child][problem_num] = 0    
        
        return ContestInfo(name=name, Tasks=problems_ids, Result=contest_result)

    

def get_contest_name(contest_id : int, niknames : Dict) -> str:
    cf = codeforces_api.CodeforcesApi(api_key=config.CF_API_KEY, secret=config.CF_API_SECRET)
    result = []
    standings = cf.contest_standings(contest_id=contest_id, handles=["VladTim"])
    result.append(standings["contest"].name)

if __name__ == "__main__":
    print("main")
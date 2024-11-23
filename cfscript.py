import codeforces_api
from typing import List, Dict
from dataclasses import dataclass

@dataclass
class ContestInfo:
    name: str
    Tasks: List[str]
    Result: List[List[int]]

class CodeforcesServer:
    def __init__(self):
        api_key = "d5b2267edde060861100e73b285fc450877b7622"
        secret = "e8d1310eca7dc3c7934614de655f49673dc14846" 
        self.cf = codeforces_api.CodeforcesApi(api_key=api_key, secret=secret)

    def generate_contest_info(self, contest_id : int, niknames : List[str]):
        standings = self.cf.contest_standings(contest_id=contest_id, handles=["VladTim"])

        name = standings["contest"].name
        problems_ids = [problem.index for problem in standings["problems"]] 
        problems_dict = {problem: idx for idx, problem in enumerate(problems_ids)}
        print(problems_ids)

        status = self.cf.contest_status(contest_id)
        contest_result = [[0 for _ in range(len(problems_ids))] for _ in range(len(niknames.keys()))]
        for row in status:
            if (row.author.members[0].handle in niknames):
                # print(row.verdict)
                # print(row.author.members[0].handle)
                # print(row.problem.index)
                num_child = niknames[row.author.members[0].handle]
                problem_num = problems_dict[row.problem.index]
                contest_result[num_child][problem_num] = 1        
        
        return ContestInfo(name=name, Tasks=problems_ids, Result=contest_result)

    

def get_contest_name(contest_id : int, niknames : Dict) -> str:
    cf = codeforces_api.CodeforcesApi(api_key=api_key, secret=secret)
    result = []
    standings = cf.contest_standings(contest_id=contest_id, handles=["VladTim"])
    result.append(standings["contest"].name)

if __name__ == "__main__":
    print("main")
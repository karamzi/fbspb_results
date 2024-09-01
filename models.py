from openpyxl import load_workbook
import json
import requests
from typing import Optional


class Player:
    name: str
    rang: str
    place: int
    games: [int]
    avg: float
    min: int
    max: int
    sum: int
    more_200: int

    def __init__(self):
        self.games = []
        self.more_200 = 0
        self.rang = ''

    def __str__(self):
        return self.name

    def to_json(self):
        self.count_statistic()
        player = {
            'name': self.name.strip(),
            'rang': self.rang if self.rang else '',
            'place': self.place,
            'games': self.games,
            'avg': self.avg,
            'min': self.min,
            'max': self.max,
            'sum': self.sum,
            'more_200': self.more_200
        }
        return player

    def count_statistic(self):
        self.games = list(filter(lambda x: x > 0, self.games))

        if len(self.games) > 0:
            self.min = min(self.games)
            self.max = max(self.games)
            self.sum = sum(self.games)
            self.avg = self.sum / len(self.games)
        else:
            self.min = 300
            self.max = 0
            self.sum = 0
            self.avg = 0
        for game in self.games:
            if game >= 200:
                self.more_200 += 1

        if self.name == 'Завертяева Елена':
            self.name = 'Завертяева Алена'


class BaseResultParse:
    man: [Player]
    woman: [Player]
    pointer: [str, str]
    file_name: str
    wb = None
    type: str

    def __init__(self, file_name: str):
        self.file_name = file_name
        self.wb = load_workbook(file_name)
        self.pointer = ['A', 1]
        self.man = []
        self.woman = []

    def change_row(self, by=1):
        self.pointer[1] += by

    def change_column(self, by=1):
        letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                   'U', 'V', 'W', 'X', 'Y', 'Z']
        current_index = letters.index(self.pointer[0])
        next_index = (current_index + by) % len(letters)
        self.pointer[0] = letters[next_index]

    def next_line(self, column='A', rows=1):
        self.pointer[0] = column
        self.pointer[1] += rows

    def get_cell_address(self):
        return f'{self.pointer[0]}{self.pointer[1]}'

    def find_player(self, target: [Player], name: str) -> Optional[Player]:
        for item in target:
            if item.name == name:
                return item
        return None

    def send_result(self, tournament_id):
        from main import DEBUG

        data = {
            'id': tournament_id,
            'type': self.type,
            'man': [],
            'woman': []
        }
        for man in self.man:
            data['man'].append(man.to_json())
        for woman in self.woman:
            data['woman'].append(woman.to_json())
        data = json.dumps(data, ensure_ascii=True)

        if DEBUG:
            response = requests.post('http://127.0.0.1:8000/api/parse_results/', data=data)
        else:
            response = requests.post('https://фбспб.рф/api/parse_results/', data=data)

        print(response.status_code)

        if response.status_code != 200:
            with open('index.html', 'w') as f:
                f.write(response.content.decode('utf-8'))
        else:
            print(json.loads(response.content))

    def parse(self):
        raise NotImplementedError

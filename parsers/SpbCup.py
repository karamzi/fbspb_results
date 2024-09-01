from models import Player, BaseResultParse


class SpbCup(BaseResultParse):
    type = 'spb_cup'

    def parse_qualification(self, target: [Player], ws):
        while not isinstance(ws[self.get_cell_address()].value, int):
            self.change_row()
        while ws[self.get_cell_address()].value is not None:
            player = Player()
            player.place = ws[self.get_cell_address()].value
            self.change_column()
            player.rang = ws[self.get_cell_address()].value
            self.change_column()
            player.name = ws[self.get_cell_address()].value
            self.change_column()
            for _ in range(6):
                game = ws[self.get_cell_address()].value
                player.games.append(game)
                self.change_column()
            target.append(player)
            self.next_line()

    def parse_finals(self, target: [Player], ws, column):
        while not isinstance(ws[self.get_cell_address()].value, int):
            self.change_row()
        last_row = None
        while ws[self.get_cell_address()].value is not None:
            count = 0
            self.change_column()
            name = ws[self.get_cell_address()].value
            player: Player = self.find_player(target, name)
            if player is None:
                raise Exception(f'Player: {name} have not been found')
            self.change_column()
            player.games.append(ws[self.get_cell_address()].value)
            self.change_column()
            player.games.append(ws[self.get_cell_address()].value)
            self.next_line(column=column)
            last_row = self.pointer[1]
            while not isinstance(ws[self.get_cell_address()].value, int) and count != 20:
                self.change_row()
                count += 1
        self.pointer = [column, last_row]
        self.change_column()
        while not isinstance(ws[self.get_cell_address()].value, str):
            self.change_row()
        while ws[self.get_cell_address()].value is not None:
            name = ws[self.get_cell_address()].value
            player: Player = self.find_player(target, name)
            if player is None:
                raise Exception(f'Player: {name} have not been found')
            self.change_column()
            place: str = ws[self.get_cell_address()].value
            place = place.strip()
            player.place = int(place.split(' ')[0])
            self.next_line(column=column)
            self.change_column()

    def parse(self):
        ws = self.wb[self.wb.sheetnames[0]]
        self.parse_qualification(self.man, ws)
        self.parse_qualification(self.woman, ws)
        self.pointer = ['A', 1]
        ws = self.wb[self.wb.sheetnames[1]]
        self.parse_finals(self.man, ws, 'A')
        self.pointer = ['G', 1]
        self.parse_finals(self.woman, ws, 'G')

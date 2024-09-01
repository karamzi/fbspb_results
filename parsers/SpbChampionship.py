from models import Player, BaseResultParse


class SpbChampionship(BaseResultParse):
    type = 'spb_championship'

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

            self.change_column(by=2)

            for _ in range(6):
                game = ws[self.get_cell_address()].value
                player.games.append(game)
                self.change_column()

            target.append(player)
            self.next_line()

    def parse_sami_finals(self, target: [Player], ws):
        while not isinstance(ws[self.get_cell_address()].value, int):
            self.change_row()

        while ws[self.get_cell_address()].value is not None:
            place = ws[self.get_cell_address()].value
            self.change_column(by=2)

            name = ws[self.get_cell_address()].value
            player = self.find_player(target, name)

            if player is None:
                raise Exception(f'Player: {name} have not been found')

            player.place = place
            self.change_column(by=2)

            for _ in range(6):
                game = ws[self.get_cell_address()].value
                player.games.append(game)
                self.change_column()

            self.next_line()


    def parse_finals(self, target: [Player], ws, column):
        while not isinstance(ws[self.get_cell_address()].value, int):
            self.change_row()

        players = 0
        while ws[self.get_cell_address()].value is not None and players < 8:
            players += 1
            self.change_column(by=2)
            name = ws[self.get_cell_address()].value
            player = self.find_player(target, name)

            if player is None:
                raise Exception(f'Player: {name} have not been found')

            self.change_column()
            player.games.append(ws[self.get_cell_address()].value)
            self.change_column()
            player.place = ws[self.get_cell_address()].value
            self.next_line(column=column)

            count = 0
            while not isinstance(ws[self.get_cell_address()].value, int) and count != 15:
                self.change_row()
                count += 1


    def parse(self):
        ws = self.wb[self.wb.sheetnames[0]]
        self.parse_qualification(self.man, ws)
        self.parse_qualification(self.woman, ws)

        self.pointer = ['A', 1]
        ws = self.wb[self.wb.sheetnames[1]]
        self.parse_sami_finals(self.man, ws)
        self.parse_sami_finals(self.woman, ws)

        self.pointer = ['B', 1]
        ws = self.wb[self.wb.sheetnames[2]]
        self.parse_finals(self.man, ws, 'B')

        self.parse_finals(self.woman, ws, 'B')



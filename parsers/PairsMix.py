from models import Player, BaseResultParse


class PairsMix(BaseResultParse):
    type = 'Teams'

    def parse_sheet(self, sheet_name: str):
        self.pointer = ['B', 1]
        ws = self.wb[sheet_name]
        while ws[self.get_cell_address()].value != 1:
            self.change_row()
        count = 0
        place = ''
        while ws[self.get_cell_address()].value is not None:
            self.change_column()
            player = Player()
            player.rang = ws[self.get_cell_address()].value
            self.change_column()
            player.name = ws[self.get_cell_address()].value
            self.change_column()
            for _ in range(6):
                game = ws[self.get_cell_address()].value
                player.games.append(game)
                self.change_column()
            self.change_column(4)
            place = ws[self.get_cell_address()].value if ws[self.get_cell_address()].value else place
            player.place = place
            if count % 2 == 0:
                self.man.append(player)
            else:
                self.woman.append(player)
            self.next_line(column='B', rows=count % 2 + 1)
            count += 1

    def parse_finals(self, sheet_name: str):
        ws = self.wb[sheet_name]
        place = ''
        while ws[self.get_cell_address()].value != 1:
            self.change_row()

        while ws[self.get_cell_address()].value is not None:
            count = 0
            self.change_column(2)
            name = ws[self.get_cell_address()].value
            player = self.find_player(self.man, name)
            if player is None:
                player = self.find_player(self.woman, name)
            self.change_column()
            player.games.append(ws[self.get_cell_address()].value)
            self.change_column(2)
            place = ws[self.get_cell_address()].value if ws[self.get_cell_address()].value else place
            player.place = place
            self.next_line(column='B')
            while ws[self.get_cell_address()].value not in [1, 2] and count != 5:
                self.change_row()
                count += 1

    def parse(self):
        self.parse_sheet(self.wb.sheetnames[0])
        self.parse_finals(self.wb.sheetnames[0])

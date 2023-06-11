from models import Player, BaseResultParse


class SecondPmParser(BaseResultParse):
    type = 'SecondPm'

    def parse_sheet(self, sheet_name: str) -> [Player]:
        result = []
        ws = self.wb[sheet_name]
        while ws[self.get_cell_address()].value != 1:
            self.change_row()
        while ws[self.get_cell_address()].value is not None:
            player = Player()
            player.place = ws[self.get_cell_address()].value
            self.change_column()
            player.rang = ws[self.get_cell_address()].value
            self.change_column()
            player.name = ws[self.get_cell_address()].value
            self.change_column(by=6)
            for _ in range(6):
                game = ws[self.get_cell_address()].value
                player.games.append(game)
                self.change_column()
            result.append(player)
            self.next_line()
        return result

    def parse_finals(self, sheet_name: str, target: [Player]):
        ws = self.wb[sheet_name]
        while ws[self.get_cell_address()].value != 1:
            self.change_row()
        while ws[self.get_cell_address()].value is not None:
            place = ws[self.get_cell_address()].value
            self.change_column()

            name = ws[self.get_cell_address()].value
            player = self.find_player(target, name)

            player.place = place
            self.change_column(by=2)

            while type(ws[self.get_cell_address()].value) is int:
                game = ws[self.get_cell_address()].value
                player.games.append(game)
                self.change_column()

            self.next_line(rows=2)

    def parse(self):
        self.man = self.parse_sheet(self.wb.sheetnames[0])
        self.woman = self.parse_sheet(self.wb.sheetnames[0])
        self.pointer = ['A', 1]
        self.parse_finals(self.wb.sheetnames[1], self.man)
        self.parse_finals(self.wb.sheetnames[1], self.woman)

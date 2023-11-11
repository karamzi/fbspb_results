from models import Player, BaseResultParse


class FirstPmParser(BaseResultParse):
    type = 'FirstPm'

    def parse_sheet(self, sheet_name: str) -> [Player]:
        self.pointer = ['A', 1]
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

    def parse(self):
        self.man = self.parse_sheet('мужчины')
        self.woman = self.parse_sheet('женщины')

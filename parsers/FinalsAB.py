from models import Player, BaseResultParse


class FinalsAB(BaseResultParse):
    type = 'finals_a_b'

    def parse_sheet(self, ws, target) -> [Player]:
        self.pointer = ['A', 1]
        while not isinstance(ws[self.get_cell_address()].value, int):
            self.change_row()
        while ws[self.get_cell_address()].value is not None:
            count = 0
            self.change_column()
            name = ws[self.get_cell_address()].value
            player = self.find_player(target, name)

            if not player:
                player = Player()
                player.name = name
                target.append(player)

            self.change_column()
            player.games.append(ws[self.get_cell_address()].value)
            self.change_column()
            player.games.append(ws[self.get_cell_address()].value)
            self.change_column(by=2)
            player.place = ws[self.get_cell_address()].value
            self.next_line()
            while not isinstance(ws[self.get_cell_address()].value, int) and count != 5:
                self.change_row()
                count += 1

    def parse(self):
        ws = self.wb[self.wb.sheetnames[0]]
        self.parse_sheet(ws, self.man)
        ws = self.wb[self.wb.sheetnames[1]]
        self.parse_sheet(ws, self.woman)

from models import Player, BaseResultParse


class Pairs(BaseResultParse):
    type = 'Teams'

    def parse_sheet(self, sheet_name: str) -> [Player]:
        self.pointer = ['B', 1]
        result = []
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
            result.append(player)
            self.next_line(column='B', rows=count % 2 + 1)
            count += 1
        return result

    def parse_finals(self, target: [Player], sheet_name: str):
        ws = self.wb[sheet_name]
        count = 0
        place = ''
        while ws[self.get_cell_address()].value not in [1, 2] and count != 5:
            self.change_row()
            count += 1
        while ws[self.get_cell_address()].value is not None:
            count = 0
            self.change_column(2)
            name = ws[self.get_cell_address()].value
            player = self.find_player(target, name)
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
        self.man = self.parse_sheet(self.wb.sheetnames[0])
        self.parse_finals(self.man, self.wb.sheetnames[0])
        self.woman = self.parse_sheet(self.wb.sheetnames[1])
        self.parse_finals(self.woman, self.wb.sheetnames[1])

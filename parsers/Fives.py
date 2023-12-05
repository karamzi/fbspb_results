from models import Player, BaseResultParse


class Fives(BaseResultParse):
    type = 'Teams'

    def parse_sheet(self, sheet_name: str) -> [Player]:
        self.pointer = ['B', 1]
        result = []
        ws = self.wb[sheet_name]

        while ws[self.get_cell_address()].value != 1:
            self.change_row()

        count = 1
        place = ''

        while ws[self.get_cell_address()].value is not None:
            self.change_column()
            player = Player()

            player.rang = ws[self.get_cell_address()].value
            self.change_column()

            player.name = ws[self.get_cell_address()].value
            self.change_column(2)

            for _ in range(6):
                game = ws[self.get_cell_address()].value
                player.games.append(game)
                self.change_column()
            self.change_column(4)

            place = ws[self.get_cell_address()].value if ws[self.get_cell_address()].value else place
            player.place = place

            result.append(player)

            if count % 5 == 0:
                self.next_line(column='B', rows=2)
            else:
                self.next_line(column='B')

            count += 1

        return result

    def parse_finals(self, target: [Player], sheet_name: str):
        ws = self.wb[sheet_name]
        count = 0
        place = ''
        while ws[self.get_cell_address()].value not in [1, 2, 3] and count != 10:
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
            while ws[self.get_cell_address()].value not in [1, 2, 3] and count != 10:
                self.change_row()
                count += 1

    def parse(self):
        self.man = self.parse_sheet(self.wb.sheetnames[0])
        self.woman = self.parse_sheet(self.wb.sheetnames[1])

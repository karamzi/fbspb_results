from parsers.FirstPm import FirstPmParser
from parsers.SecondPm import SecondPmParser
from parsers.Pairs import Pairs
from parsers.PairsMix import PairsMix
from parsers.Triples import Triples
from parsers.FinalsA import FinalsA
from parsers.FinalsB import FinalsB
from parsers.SpbCup import SpbCup
from parsers.TriplesMix import TriplesMix
from parsers.Fives import Fives
from parsers.SpbChampionship import SpbChampionship

DEBUG = False

TOURNAMENTS = {
    'FirstPm': FirstPmParser,
    'SecondPm': SecondPmParser,
    'Pairs': Pairs,
    'PairsMix': PairsMix,
    'Triples': Triples,
    'TriplesMix': TriplesMix,
    'finals_a': FinalsA,
    'finals_b': FinalsB,
    'spb_cup': SpbCup,
    'Fives': Fives,
    'SpbChampionship': SpbChampionship
}


def choose_type():
    print('1 - 1 регламент ПМ (6 игр)')
    print('2 - 2 регламент ПМ (12 игр)')
    print('3 - Финалы А')
    print('4 - Пары')
    print('5 - Пары микс')
    print('6 - Тройки')
    print('7 - Кубок СПБ')
    print('8 - Тройки микс')
    print('9 - Пятерки')
    print('10 - Чемпионат СПБ')
    print('11 - Финалы Б')

    option = int(input('Выберите тип турнира: '))

    if option == 1:
        return 'FirstPm'
    elif option == 2:
        return 'SecondPm'
    elif option == 3:
        return 'finals_a'
    elif option == 4:
        return 'Pairs'
    elif option == 5:
        return 'PairsMix'
    elif option == 6:
        return 'Triples'
    elif option == 7:
        return 'spb_cup'
    elif option == 8:
        return 'TriplesMix'
    elif option == 9:
        return 'Fives'
    elif option == 10:
        return 'SpbChampionship'
    elif option == 11:
        return 'finals_b'

    raise Exception('Неверный тип')


if __name__ == '__main__':
    tournament_id = int(input('Введите id турнира: '))
    tournament_type = choose_type()
    file_name = input('Введите название файла: ')

    tournament = TOURNAMENTS[tournament_type](file_name)
    tournament.parse()
    tournament.send_result(tournament_id)

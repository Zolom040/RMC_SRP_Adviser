from typing import List

# Класс для загрузки списков валидных ID из соответствующих файлов
class esi_ID_loader:
    def __init__(self):
        self.alliances = self._load_from_file('Alliance_ids.txt')
        self.dps_ships = self._load_from_file('DpsShips_ids.txt')
        self.support_ships = self._load_from_file('SupportShips_ids.txt')
        self.tackle_ships = self._load_from_file('Tackle_ids.txt')
        self.vaiuble_ships = self._load_from_file('ValubleShip_ids.txt')
        self.capital_ships = self._load_from_file('CapitalShip_ids.txt')

    @staticmethod
    def _load_from_file(filename):
        with open(filename, 'r') as file:
            # Читаем все строки файла и сохраняем их в список
            lines = file.readlines()

        # Убираем символы новой строки у каждой строки (если они есть)
        return [line.strip() for line in lines]


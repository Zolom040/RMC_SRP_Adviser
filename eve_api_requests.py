import requests
import time

MAX_RETRIES = 3  # Максимальное количество попыток

class EveRequester:
    def __init__(self):
        print("done")

    # Запрос на получение всех потерь заданного альянса в заданный период (1 месяц)
    # Ответ -  killmail_id, zkb [location, hash, fittValue (isk), dropValue (isk), destroyedValue, totalvalue, points, npc, solo, awox, labels ]
    def get_zkillboard_killmails_for_alliance(self, alliance_id=None, year=2025, month=1, page=1):
        time.sleep(2)
        url = f"https://zkillboard.com/api/losses/allianceID/{alliance_id}/year/{year}/month/{month}/page/{page}/"
        response = requests.get(url)
        return response.json()

    # Запрос на детали килмыла (нужно чтобы определить шип
    # Ответ - attackers, killmail_id, killmail_time, solar_system_id, victim [aliance_id, character_id, corporation_id,
    # items, position [x,y,z], ship_type_id, faction_id]
    def get_killmail_details(self, killmail_id, killmail_hash):
        url = f"https://esi.evetech.net/latest/killmails/{killmail_id}/{killmail_hash}/"
        retries_left = MAX_RETRIES

        while retries_left > 0:
            try:
                response = requests.get(url, timeout=5)

                if response.status_code == 200:
                    response_json = response.json()

                    if "error" not in response_json:
                        return response_json
                    else:
                        raise ValueError("Response contains an error")

                else:
                    print(f"Ошибка: {response.status_code}")
                    retries_left -= 1
                    time.sleep(60)  # Ждем перед новой попыткой
                    continue

            except Exception as e:
                print(f"Ошибка при получении данных: {e}")
                retries_left -= 1
                time.sleep(60)  # Ждем перед новой попыткой
                continue

        # Если мы вышли из цикла, значит исчерпаны все попытки
        print("Все попытки исчерпаны, не удалось получить данные.")
        return None

    # Запрос на инфу о персонаже нужно только name
    def get_CharacterInfo(self, character_id):
        url = f"https://esi.evetech.net/dev/characters/{character_id}/"
        retries_left = MAX_RETRIES

        while retries_left > 0:
            try:
                response = requests.get(url, timeout=5)

                if response.status_code == 200:
                    response_json = response.json()

                    if "error" not in response_json:
                        return response_json
                    else:
                        raise ValueError("Response contains an error")

                else:
                    print(f"Ошибка: {response.status_code}")
                    retries_left -= 1
                    time.sleep(60)  # Ждем перед новой попыткой
                    continue

            except Exception as e:
                print(f"Ошибка при получении данных: {e}")
                retries_left -= 1
                time.sleep(60)  # Ждем перед новой попыткой
                continue

        # Если мы вышли из цикла, значит исчерпаны все попытки
        print("Все попытки исчерпаны, не удалось получить данные.")
        return None
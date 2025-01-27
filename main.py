

from eve_api_requests import EveRequester
from esi_IDLoader import esi_ID_loader
from datetime import datetime, timedelta, date
import pandas as pd
import json
import openpyxl
import time

def check_field(json, field_name):
    if field_name in json:
        return json[field_name]
    else:
        print(f"Поля `{field_name}` нет в JSON.")

def main():
    loader = esi_ID_loader()
    current_date = datetime.now()
    last_month = date.today().replace(day=1) - timedelta(days=1)
    month = last_month.month
    year = last_month.year
    print(f"month {month} year {year}")
    requester = EveRequester()
    Members = dict()
    for aliance in loader.alliances:
        for p in range(1,15):
            aliance_killmails = requester.get_zkillboard_killmails_for_alliance(aliance,year,month,p)
            for km in aliance_killmails:
                killmail_id = km["killmail_id"]
                killmail_hash = km["zkb"]["hash"]
                totalValue = km["zkb"]["totalValue"]
                kmd = requester.get_killmail_details(killmail_id, killmail_hash)

                if "victim" in kmd:
                    if "character_id" in kmd["victim"]:
                        char_id = kmd["victim"]["character_id"]
                    else:
                        print(f"какая то ошибка с json в киле {killmail_id} пропущу это килмыло. ктото недополучит компенс")
                        continue
                else:
                    print(f"какая то ошибка с json в киле {killmail_id} пропущу это килмыло. ктото недополучит компенс")
                    continue

                if char_id not in Members:
                    new_member = {
                        "ID": char_id,
                        "Name": requester.get_CharacterInfo(kmd["victim"]["character_id"])["name"],
                        "total_cost": 0,
                        "natur_compens_ids" : []
                    }
                    Members[new_member["ID"]] = new_member

                sh_id = kmd["victim"]["ship_type_id"]
                ship_id = str(sh_id)
                if ship_id in loader.support_ships:
                    Members[char_id]["total_cost"] += totalValue
                    print(Members[char_id]["Name"] + " Большой молодец ")
                elif ship_id in loader.tackle_ships:
                    Members[char_id]["total_cost"] += totalValue
                    print(Members[char_id]["Name"] + " молодец ")
                elif ship_id in loader.dps_ships:
                    Members[char_id]["total_cost"] += totalValue * 0.7
                    print(Members[char_id]["Name"] + " норм ")
                elif ship_id in loader.vaiuble_ships:
                    Members[char_id]["natur_compens_ids"].append(ship_id)
                    print(Members[char_id]["Name"] + " молодец ")
                elif ship_id in loader.capital_ships:
                    Members[char_id]["natur_compens_ids"].append(ship_id)
                    print(Members[char_id]["Name"] + " Вообще молодец ")
                else:
                    print(Members[char_id]["Name"] + " крабина ")
            time.sleep(2)

    # Фильтрация словаря
    filtered_data = {k: v for k, v in Members.items()}

    # Исключаем из словаря все элементы, у которых total_cost и natur_compens_ids пустые
    filtered_data = {k: v for k, v in filtered_data.items() if v["total_cost"] > 0 or len(v["natur_compens_ids"]) > 0}

    df = pd.DataFrame(list(filtered_data.values()))
    timestamp = current_date.strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"output_{timestamp}.xlsx"
    df.to_excel(filename, index=False)

    print(f"Excel-файл '{filename}' успешно создан!")











if __name__ == "__main__":
    main()
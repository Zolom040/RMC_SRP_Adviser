

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
    inshu = requester.get_Inshurances()
    inshurances = dict()
    for ins in inshu:
        for level in ins["levels"]:
            if level["name"] == "Platinum":
                inshurances[ins["type_id"]] = level["payout"]

    Members = dict()
    for aliance in loader.alliances:
        AliInfo = requester.get_Aliance_Info(aliance)
        AlianceTicker = AliInfo["ticker"]
        for p in range(1,15):
            aliance_killmails = requester.get_zkillboard_killmails_for_alliance(aliance,year,month,p)
            for km in aliance_killmails:
                killmail_id = km["killmail_id"]
                killmail_hash = km["zkb"]["hash"]
                totalValue = km["zkb"]["fittedValue"]
                cost = 0
                if not km["zkb"]["npc"]:
                    kmd = requester.get_killmail_details(killmail_id, killmail_hash)
                    if "victim" in kmd:
                        if "character_id" in kmd["victim"]:
                            char_id = kmd["victim"]["character_id"]
                        else:
                            print(f"какая то ошибка с json в киле {killmail_id} пропущу это килмыло. ктото недополучит компенс. вероятно это мобилка")
                            continue
                    else:
                        print(f"какая то ошибка с json в киле {killmail_id} пропущу это килмыло. ктото недополучит компенс. вероятно это мобилка")
                        continue

                    if char_id not in Members:
                        new_member = {
                            "ID": char_id,
                            "Name": requester.get_CharacterInfo(kmd["victim"]["character_id"])["name"],
                            "Aliance": AlianceTicker,
                            "total_cost": 0,
                            "natur_compens_ids" : [],
                            "dps_links": [],
                            "support_links": [],
                            "tackle_links": [],
                            "valuble_links": [],
                            "capital_links": []
                        }
                        Members[new_member["ID"]] = new_member

                    sh_id = kmd["victim"]["ship_type_id"]
                    ship_id = str(sh_id)
                    link = f"https://zkillboard.com/kill/{killmail_id}/"
                    cost = 0
                    # forbidden = [33475, 33700]
                    if sh_id in inshurances:
                        cost = totalValue - inshurances[sh_id]

                    if ship_id in loader.support_ships:
                        Members[char_id]["total_cost"] += cost
                        Members[char_id]["support_links"].append({"cost": cost, "link": link})
                        print(Members[char_id]["Name"] + " Большой молодец ")
                    elif ship_id in loader.tackle_ships:
                        Members[char_id]["total_cost"] += cost
                        Members[char_id]["tackle_links"].append({"cost": cost, "link": link})
                        print(Members[char_id]["Name"] + " молодец ")
                    elif ship_id in loader.dps_ships:
                        Members[char_id]["total_cost"] += cost * 0.7
                        Members[char_id]["dps_links"].append({"cost": cost, "link": link})
                        print(Members[char_id]["Name"] + " норм ")
                    elif ship_id in loader.vaiuble_ships:
                        Members[char_id]["natur_compens_ids"].append(ship_id)
                        Members[char_id]["total_cost"] -= inshurances[sh_id]
                        cost = -inshurances[sh_id]
                        Members[char_id]["valuble_links"].append({"cost": cost, "link": link})
                        print(Members[char_id]["Name"] + " молодец ")
                    elif ship_id in loader.capital_ships:
                        Members[char_id]["natur_compens_ids"].append(ship_id)
                        Members[char_id]["total_cost"] -= inshurances[sh_id]
                        cost = -inshurances[sh_id]
                        Members[char_id]["capital_links"].append({"cost": cost, "link": link})
                        print(Members[char_id]["Name"] + " Вообще молодец ")
                    else:
                        print(Members[char_id]["Name"] + " крабина ")
                else:
                    print("Килл от нпц")
            time.sleep(2)

    MData = dict()
    for char_id, member_data in Members.items():
        if member_data["total_cost"] != 0:
            main_data = {
                "ID": member_data["ID"],
                "Name": member_data["Name"],
                "Aliance": member_data["Aliance"],
                "total_cost": member_data["total_cost"],
                "natur_compens_ids": member_data["natur_compens_ids"]
            }
            MData[char_id] = main_data

    main_df = pd.DataFrame(list(MData.values()))
    links_data = []
    # Проходимся по каждому члену в Members
    for char_id, member_data in Members.items():
        # Проверяем, что total_cost != 0
        if member_data["total_cost"] != 0:
            # Проходимся по всем типам линков
            for link_type, links in member_data.items():
                if "links" in link_type and links:  # Проверяем, что это список линков и он не пустой
                    for link in links:
                        # Добавляем данные в links_data
                        links_data.append({
                            "Name": member_data["Name"],
                            "Aliance": member_data["Aliance"],
                            "Link Type": link_type,
                            "Cost": link["cost"],
                            "Link": link["link"]
                        })

    links_df = pd.DataFrame(links_data)
    # Сортируем DataFrame по столбцу "Name" (по алфавиту)
    links_df = links_df.sort_values(by=["Aliance", "Name"])

    # # Фильтрация словаря
    # filtered_data = {k: v for k, v in Members.items()}
    #
    # # Исключаем из словаря все элементы, у которых total_cost и natur_compens_ids пустые
    # filtered_data = {k: v for k, v in filtered_data.items() if v["total_cost"] > 0 or len(v["natur_compens_ids"]) > 0}
    #
    # df = pd.DataFrame(list(filtered_data.values()))
    timestamp = current_date.strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"output_{timestamp}.xlsx"
    # df.to_excel(filename, index=False)

    with pd.ExcelWriter(filename) as writer:
        main_df.to_excel(writer, sheet_name="Main", index=False)
        links_df.to_excel(writer, sheet_name="Links for check", index=False)

    print(f"Excel-файл '{filename}' успешно создан!")











if __name__ == "__main__":
    main()
from typing import Optional
import json
import pandas
from rich import print
from rich.table import Table
import sys
import re


def get_address_number(street: str, address: str) -> Optional[int]:
    try:
        number = re.search(rf"{street}(\d+)号", address)
        if number:
            return int(number.group(1))
        else:
            number = re.search(rf"{street}(.+)号", address).group(1)
            return int(re.search(r"\d+", number).group(0))
    except:
        print(f"[red]Error: {street} {address}")
        return


def find_neighborhood(address: str, streets: dict[str, str]) -> Optional[str]:
    neighborhood_name = None
    for street, neighborhoods in streets.items():
        if street in address:
            for neighborhood in neighborhoods:
                if neighborhood["all"]:
                    neighborhood_name = neighborhood["name"]
                    break
                else:
                    address_number = get_address_number(street, address)
                    if neighborhood["oddity"] == address_number % 2:
                        if (
                            neighborhood["start"]
                            <= address_number
                            <= neighborhood["end"]
                        ):
                            neighborhood_name = neighborhood["name"]
                            break
            break
    return neighborhood_name


if __name__ == "__main__":
    streets = json.load(open("streets.json", encoding="utf-8"))

    neighborhood_col_name = "所属居委"
    neighborhood_column = []

    table = Table()
    table.add_column("地址", style="green", no_wrap=True)
    table.add_column(neighborhood_col_name, style="blue")

    filename, address_column = sys.argv[1], sys.argv[2]
    excel = pandas.read_excel(filename)
    addresses = excel[address_column].values

    for address in addresses:
        neighborhood_name = find_neighborhood(address, streets)
        if neighborhood_name:
            table.add_row(address, neighborhood_name)
            neighborhood_column.append(neighborhood_name)
        else:
            table.add_row(address, "[red]未知")
            neighborhood_column.append("未知")

    print(table)
    if neighborhood_col_name not in excel:
        excel.insert(len(excel.keys()), neighborhood_col_name, neighborhood_column)
    else:
        excel[neighborhood_col_name] = neighborhood_column
    excel.sort_values(
        by=neighborhood_col_name, ascending=True, inplace=True, ignore_index=True
    )
    excel.to_excel(filename, index=False)

import os
import csv
from src.excel_file_cp import create_xlsx
from src.excel_edit import refresh_rows_have_data
from dotenv import load_dotenv

load_dotenv()
SHEET_NAME2 = os.getenv("SHEET_NAME2")
SHEET_NAME3 = os.getenv("SHEET_NAME3")

def read_input():
    ret = []
    keys = ["name", "repository_name", "branch_name"]
    try:
        with open("input.csv", "r") as f:
            csv_reader = csv.reader(f)
            for row in csv_reader:
                ret.append(dict(zip(keys, [val.strip() for val in row])))
        return ret
    except FileNotFoundError:
        print("csv file not found")
        exit()

if __name__ == '__main__':
    csv_read = read_input()
    # create copies of original xlsx
    create_xlsx([val["name"] for val in csv_read])
    # edit every copy files
    for vals in csv_read:
        # 1. delete sheet2 & sheet3
        refresh_rows_have_data(f".xlsx", SHEET_NAME2, )

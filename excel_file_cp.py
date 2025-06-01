import csv
import os
import shutil
from dotenv import load_dotenv

load_dotenv()
ORIGINAL_EXCEL_PATH = os.getenv("ORIGINAL_EXCEL_PATH")
FOLDER_PATH = os.getenv("FOLDER_PATH")

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

def xlsx_copy(source_path, dest_path):
    try:
        # copy file
        shutil.copy(source_path, dest_path)
    except FileNotFoundError:
        print("original file not found")
        exit()
    except Exception as e:
        print(e)
        exit()

def xlsx_rename(old_path, new_name):
    try:
        new_path = os.path.join(FOLDER_PATH, new_name)
        os.rename(old_path, new_path)
    except FileNotFoundError:
        print("rename file not found")
        exit()
    except Exception as e:
        print(e)
        exit()

def create_xlsx():
    csv_read = read_input()

    if not os.path.exists(FOLDER_PATH):
        os.makedirs(FOLDER_PATH)

    for num, row in enumerate(csv_read, start=1):
        tmp_path = os.path.join(FOLDER_PATH, f"tmp_{num}.xlsx")
        xlsx_copy(ORIGINAL_EXCEL_PATH, tmp_path)
        xlsx_rename(tmp_path, f"{num}_{row["name"]}.xlsx")

if __name__ == '__main__':
    create_xlsx()
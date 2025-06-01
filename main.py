import csv
from src.excel_file_cp import create_xlsx

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
    create_xlsx([val["name"] for val in csv_read])

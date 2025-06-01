import os
import shutil
from dotenv import load_dotenv

load_dotenv()
ORIGINAL_EXCEL_PATH = os.getenv("ORIGINAL_EXCEL_PATH")
FOLDER_PATH = os.getenv("FOLDER_PATH")

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

def create_xlsx(names):
    if not os.path.exists(FOLDER_PATH):
        os.makedirs(FOLDER_PATH)

    for num, name in enumerate(names, start=1):
        tmp_path = os.path.join(FOLDER_PATH, f"tmp_{num}.xlsx")
        xlsx_copy(ORIGINAL_EXCEL_PATH, tmp_path)
        xlsx_rename(tmp_path, f"{num}_{name}.xlsx")

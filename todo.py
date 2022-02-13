from openpyxl import load_workbook
from datetime import datetime
import pandas as pd
import os

# excel path
path = os.getenv('APPDATA') + '/vermera/todo.xlsx'


def insert_todo(todo):
    workbook = load_workbook(filename=path)
    spreadsheet = workbook.active
    spreadsheet.insert_rows(idx=2)
    spreadsheet["A2"] = todo
    spreadsheet["B2"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    spreadsheet["C2"] = "Pending"
    workbook.save(filename=path)


# todo show only pending tasks

def show_todo():
    data = pd.read_excel(path, sheet_name='Sheet')
    print(data)

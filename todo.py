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
    data = data.loc[data['status'] == 'Pending']
    print(data)


def finish_todo(identifier):
    workbook = load_workbook(filename=path)
    spreadsheet = workbook.active
    new_id = int(identifier) + 2
    if spreadsheet["C"+str(new_id)].value is not None:
        spreadsheet["C"+str(new_id)] = "Completed"
    workbook.save(filename=path)

def show_history():
    data = pd.read_excel(path, sheet_name='Sheet')
    print(data)

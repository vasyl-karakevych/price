from AGD import AGD
import os
from openpyxl import load_workbook

def exists(path):
    try:
        os.stat(path)
    except OSError:
        # print(f"File: {path} isn`t exists")
        return False
    print(f"File: {path} is exists")
    return True

def FromMarpa(agd):
    if exists('MARPA.xlsx'):
        wb = load_workbook('MARPA.xlsx')
        sheet = wb.active  
        # agd.append(name = sheet.cell(column = 1, row = 2).value, 
        #            count = sheet.cell(column = 2, row = 2).value)
        obj = AGD(name = sheet.cell(column = 1, row = 2).value, 
                    count = sheet.cell(column = 2, row = 2).value,
                    rezervacion=sheet.cell(column = 3, row = 2).value,
                    description=sheet.cell(column = 4, row = 2).value,
                    price=sheet.cell(column = 5, row = 2).value,
                    country=sheet.cell(column = 6, row = 2).value)
    agd.append(obj)
    return agd

from AGD import AGD
import os
from openpyxl import load_workbook

# exists file
def exists(path):
    try:
        os.stat(path)
    except OSError:
        print(f"File: {path} isn`t exists")
        return False
    # print(f"File: {path} is exists")
    return True
#convert fild PRICE in normal number

def FromBiss(agd):
    if exists('BISS/BISS.xlsx'):
        wb = load_workbook('BISS/BISS.xlsx')
        sheet = wb.active  

        # print("max row in BISS =" + str(sheet.max_row))
        for l in range(2, int(sheet.max_row)+1):
            # print(f"try add {l}")
            obj = AGD(name = sheet.cell(column = 1, row = l).value, 
                        count = sheet.cell(column = 4, row = l).value,
                        rezervacion = sheet.cell(column = 3, row = l).value,
                        price = int(sheet.cell(column = 6, row = l).value),
                        sklad = 'biss')
            agd.append(obj)
    return agd

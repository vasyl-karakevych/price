from AGD import AGD
import os
from openpyxl import load_workbook

# exists file
def exists(path):
    try:
        os.stat(path)
    except OSError:
        # print(f"File: {path} isn`t exists")
        return False
    print(f"File: {path} is exists")
    return True
#convert fild PRICE in normal number
def PriceToBrutto(price):
    if price[0] == '.': return 0
    price = price.replace("  ", " ")
    price = price.split(' ')
    # print (price)
    # if len(price) > 1:
    #     if price[1][0] == 'B': price = int(price[0])
    #     elif price[1][0] == 'N': price = round(int(price[0])*1.23)
    #     else: price = 0
    # else: price = 0
    return price[0]

def FromMarpa(agd):
    if exists('MARPA/MARPA.xlsx'):
        wb = load_workbook('MARPA/MARPA.xlsx')
        sheet = wb.active  

        # print("max row in MARPA =" + str(sheet.max_row))
        for l in range(2, int(sheet.max_row)+1):
            # print(f"try add {l}")
            obj = AGD(name = sheet.cell(column = 1, row = l).value, 
                        count = sheet.cell(column = 2, row = l).value,
                        rezervacion = sheet.cell(column = 3, row = l).value,
                        description = sheet.cell(column = 4, row = l).value,
                        price = int(PriceToBrutto(sheet.cell(column = 5, row = l).value)),
                        country = sheet.cell(column = 6, row = l).value,
                        sklad = 'marpa')
            agd.append(obj)
    return agd

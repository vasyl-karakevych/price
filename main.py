from itertools import count
from MARPA.marpa import FromMarpa, exists
from AGD import AGD
from openpyxl import load_workbook
# from notebooks.notebooks import WriteToNotebooks

agd = []
delivery = 1.2
NORMA = 2300


#print object to console
def PrintAGD(obj):
    print(f"name= {obj.GetName()}\n\
            type= {obj.type}\n\
            count= {obj.count}\n\
            rezervacion= {obj.rezervacion}\n\
            description= {obj.description}\n\
            country= {obj.country}\n\
            price= {obj.price}\n\
            price_with_delivery= {obj.price_with_delivery}\n")

# count and add delivery to price
def PriceDelivery():
    for obj in agd:
        if   obj.name.find("CHŁODZ") >= 0: obj.type = "freezer"
        elif obj.name.find("PRALKA") >= 0: obj.type = "pralka"
    
    for obj in agd:
        percent = (obj.price - NORMA)/1.23*(delivery-1)
        
        if obj.type == "freezer": 
            if obj.price > 1000 and obj.price < NORMA: 
                obj.price_with_delivery = round(obj.price/1.23+450)
            elif obj.price > NORMA:
                obj.price_with_delivery = round(obj.price/1.23+450+percent)
        elif obj.type == "pralka":
            if obj.price < NORMA: 
                obj.price_with_delivery = round(obj.price/1.23+400)
            elif obj.price > NORMA:              
                obj.price_with_delivery = round(obj.price/1.23+400+percent)
        else: obj.price_with_delivery = round(obj.price/1.23*delivery)

#load KURS PLN
# def LoadKursPLN():
def WriteToPrice():
    wb = load_workbook('price_vasyl.xlsx')
    sheet = wb.active
    
    # Clear
    maxColumn = int(sheet.max_column + 1)
    maxRow = int(sheet.max_row + 1)
    for i in range(1, maxColumn):
        for j in range(2, maxRow):
            sheet.cell(column=i, row=j).value = None 

    counter = 1
    for obj in agd:
        sheet.cell(column=1, row=counter+1).value = counter
        sheet.cell(column=2, row=counter+1).value = obj.name
        sheet.cell(column=3, row=counter+1).value = obj.description
        sheet.cell(column=4, row=counter+1).value = obj.count
        sheet.cell(column=5, row=counter+1).value = obj.rezervacion
        sheet.cell(column=6, row=counter+1).value = obj.price
        sheet.cell(column=7, row=counter+1).value = obj.price_with_delivery
        formula_uah=f"=G{counter+1}*$K$1"
        sheet.cell(column=8, row=counter+1).value = formula_uah
        sheet.cell(column=9, row=counter+1).value = "marpa"
        sheet.cell(column=10, row=counter+1).value = obj.type
        counter += 1
    print(f"Write to price: {counter-1} objects")
    wb.save("price_vasyl.xlsx") 


# Load from MARPA and write to price
FromMarpa(agd)
print(f"Load from MARPA: {len(agd)} objects")
PriceDelivery()
WriteToPrice()

# LAPTOPS
# laptops = AGD()
# WriteToNotebooks(laptops)
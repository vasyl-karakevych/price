from itertools import count
from marpa import FromMarpa
from AGD import AGD

agd = []
delivery = 1.2
NORMA = 2300

def PrintAGD(obj):
    print(f"name= {obj.name}\n\
            type= {obj.type}\n\
            count= {obj.count}\n\
            rezervacion= {obj.rezervacion}\n\
            description= {obj.description}\n\
            country= {obj.country}\n\
            price= {obj.price}\n\
            price_with_delivery= {obj.price_with_delivery}\n")

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

FromMarpa(agd)
print(f"Load from MARPA: {len(agd)} objects")
PriceDelivery()
# for i in agd:
#     PrintAGD(i)
PrintAGD(agd[162])

    # wb.save('MARPA.xlsx')


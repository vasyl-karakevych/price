from itertools import count
from marpa import FromMarpa
from AGD import AGD
# exists file

agd = []
delivery = 1.2

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
        obj.price_with_delivery = round(obj.price/1.23*delivery)

FromMarpa(agd)
print(f"Load from MARPA: {len(agd)} objects")
PriceDelivery()
# for i in agd:
#     PrintAGD(i)
PrintAGD(agd[0])

    # wb.save('MARPA.xlsx')


from itertools import count
from marpa import FromMarpa
from AGD import AGD
# exists file

agd = []
def PrintAGD(obj):
    print(f"name= {obj.name}\n\
            count= {obj.count}\n\
            rezervacion= {obj.rezervacion}\n\
            description= {obj.description}\n\
            country= {obj.country}\n\
            price= {obj.price}\n")


FromMarpa(agd)
PrintAGD(agd[0])


    # wb.save('MARPA.xlsx')


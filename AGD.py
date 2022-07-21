from itertools import count


class AGD:
    name = ""
    count = 0
    rezervacion = 0
    description = ""
    country = ""
    price = 0

    def __init__(self, name="", count=0, rezervacion=0, description="", country="", price=0):
        self.name = name
        self.count = count
        self.rezervacion = rezervacion
        self.description = description
        self.country = country
        self.price = price

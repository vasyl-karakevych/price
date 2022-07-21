# class AGD
class AGD:
    name = ""
    type = ""
    count = 0
    rezervacion = 0
    description = ""
    country = ""
    price = 0
    price_with_delivery = 0

    def __init__(self, name="", type="agd", count=0, rezervacion=0, description="", 
                        country="", price=0, price_with_delivery=0):
        self.name = name
        self.type = type
        self.count = count
        self.rezervacion = rezervacion
        self.description = description
        self.country = country
        self.price = price
        self.price_with_delivery = price_with_delivery

    def GetName(self): return self.name
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
    sklad = ""

    def __init__(self, name="", type="agd", count=0, rezervacion=0, description="", 
                        country="", price=0, price_with_delivery=0, sklad=""):
        self.name = name
        self.type = type
        self.count = count
        self.rezervacion = rezervacion
        self.description = description
        self.country = country
        self.price = price
        self.price_with_delivery = price_with_delivery
        self.sklad = sklad

    # return values of AGD
    def GetName(self): return self.name
    def GetType(self): return self.type
    def GetCount(self): return self.count
    def GetRezervacion(self): return self.rezervacion
    def GetDescription(self): return self.description
    def GetCountry(self): return self.country
    def GetPrice(self): return self.price
    def GetPriceWD(self): return self.price_with_delivery
    def GetSklad(self): return self.sklad

    # set values of AGD
    def SetName(self, name):
        self.name = name
        # if isinstance(name, str): self.name = name
        # else print(f"Name ={name}. The value is not string!!!")
    
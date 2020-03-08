class Tilaus:
    def __init__(self,asiakasnumero,rullakot = 0,eurolavat = 0,finlavat = 0):
        self.asiakasnumero = asiakasnumero
        self.rullakot = rullakot
        self.eurolavat = eurolavat
        self.finlavat = finlavat

    def __str__(self):
        return "Asiakasnumero: {}, Rullakoita: {}, Eurolavoja: {}, Finlavoja: {}".format(self.asiakasnumero,self.rullakot,self.eurolavat,self.finlavat)

    def onkoPakkasia(self):
        onkoPakasteita= False
        pakasteet = self.rullakot + self.eurolavat + self.finlavat
        if(pakasteet >0):
            onkoPakasteita = True
        return onkoPakasteita
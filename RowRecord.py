class RowRecord:
    id = ""
    nazwa = ""
    faktura = ""
    winType = ""
    officeType = ""

    def __init__(self, id, nazwa, faktura):
        self.id = id
        self.nazwa = nazwa
        self.faktura = faktura

    def showRowObject(self):
        print("Id", self.id, " ", "Nazwa", self.nazwa, " ", "Faktura", self.faktura)


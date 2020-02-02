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
        print("Id")
        print(self.id)
        print("Nazwa")
        print(self.nazwa)
        print("Faktura")
        print(self.faktura)


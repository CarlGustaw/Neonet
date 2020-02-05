class RowRecord:
    idDKF = ""
    nazwa = ""
    winType = ""
    officeType = ""

    def __init__(self, nazwa, idDKF):
        self.nazwa = nazwa
        self.idDKF = idDKF

    def showRowObject(self):
        print("Nazwa", self.nazwa, " ", "ID-DKF", self.idDKF)

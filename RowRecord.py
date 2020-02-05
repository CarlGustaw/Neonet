class RowRecord:
    idDKF = ""
    name = ""
    winType = ""
    officeType = ""

    def __init__(self, name, idDKF):
        self.name = name
        self.idDKF = idDKF

    def showRowObject(self):
        print("Nazwa", self.name, " ", "ID-DKF", self.idDKF)

class RowRecord:
    idDKF = ""

    def __init__(self, idDKF):
        self.idDKF = idDKF

    def showRowObject(self):
        print("ID-DKF: ", self.idDKF)

    def getID_DKF(self):
        return self.idDKF

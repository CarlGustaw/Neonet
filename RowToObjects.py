import ExcelReader
from RowRecord import RowRecord


class RowMaker:
    dataSheet = ""
    listOfObjects = []
    rowsWithBadDKF = []

    def __readExcelFile(self, mainExcelPathName):
        self.dataSheet = ExcelReader.readExcelFile(mainExcelPathName)
        print("Main excel read successfully")

    def readExcelFileToSheetAndMakingObject(self, mainExcelPathName):
        self.__readExcelFile(mainExcelPathName)

        for rowNumber in range(1, self.dataSheet.nrows - 1):

            # Check if row contain incorrect DKF == bad list else add good DKFs to correct list
            if str(self.dataSheet.row_slice(rowNumber)[8].value).find("+") != -1 or \
                    self.dataSheet.row_slice(rowNumber)[8].value == "" or \
                    self.dataSheet.row_slice(rowNumber)[8].value == "brak skanu" or \
                    self.dataSheet.row_slice(rowNumber)[8].value == "FS-01WW/00032360/2014 - ZŚW 799/2014" or \
                    self.dataSheet.row_slice(rowNumber)[8].value == "ID:379199":
                self.rowsWithBadDKF.append(rowNumber)
            else:
                self.listOfObjects.append(
                    RowRecord(self.dataSheet.row_slice(rowNumber)[3], self.dataSheet.row_slice(rowNumber)[8].value))

        print("Rows as objects add successfully", " Number of readed rows: ", self.dataSheet.nrows - 1)
        return self.listOfObjects, self.rowsWithBadDKF

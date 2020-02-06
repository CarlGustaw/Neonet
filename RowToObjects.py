import ExcelReader
from RowRecord import RowRecord


class RowMaker:
    dataSheet = ""

    @staticmethod
    def readExcelFileToSheetAndMakingObject(mainExcelPathName):
        dataSheet = ExcelReader.readExcelFile(mainExcelPathName)
        print("Main excel read successfully")
        listOfObjects = []
        rowsWithBadDKF = []

        for rowNumber in range(1, dataSheet.nrows - 1):

            # Check if row contain correct DKF
            if str(dataSheet.row_slice(rowNumber)[8].value).find("+") != -1 or \
                    dataSheet.row_slice(rowNumber)[8].value == "" or \
                    dataSheet.row_slice(rowNumber)[8].value == "brak skanu" or \
                    dataSheet.row_slice(rowNumber)[8].value == "FS-01WW/00032360/2014 - ZŚW 799/2014" or \
                    dataSheet.row_slice(rowNumber)[8].value == "ID:379199":
                rowsWithBadDKF.append(rowNumber)
            else:
                listOfObjects.append(RowRecord(dataSheet.row_slice(rowNumber)[3], dataSheet.row_slice(rowNumber)[8].value))

        print("Rows as objects add successfully", " Number of readed rows: ", dataSheet.nrows - 1)
        print()
        return listOfObjects, rowsWithBadDKF

import ExcelReader
from RowRecord import RowRecord


class RowMaker:
    dataSheet = ""

    @staticmethod
    def readExcelFileToSheetAndMakingObject(excelPathName):
        dataSheet = ExcelReader.readExcelFile(excelPathName)
        print("Excel read successfully")
        listOfObjects = []

        for rowNumber in range(1, dataSheet.nrows - 1):
            listOfObjects.append(RowRecord(dataSheet.row_slice(rowNumber)[3], dataSheet.row_slice(rowNumber)[8]))

        print("Rows as objects add successfully", " Number of readed rows: ", dataSheet.nrows - 1)
        return listOfObjects

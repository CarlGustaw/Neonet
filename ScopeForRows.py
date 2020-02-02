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
            print(rowNumber)
            listOfObjects.append(RowRecord(dataSheet.row_slice(rowNumber)[0], dataSheet.row_slice(rowNumber)[1],
                                                dataSheet.row_slice(rowNumber)[3]))

        for rowObject in listOfObjects:
            rowObject.showRowObject()

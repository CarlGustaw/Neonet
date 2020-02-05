import ExcelReader


class SearchEngineExcel:
    dataSheet = ""

    @staticmethod
    def ScanFileForPatterns(excelPathName):
        dataSheet = ExcelReader.readExcelFile(excelPathName)

        for rowNumber in range(0, dataSheet.nrows - 1):
            print(dataSheet.row_slice(rowNumber))

import ExcelReader


class SearchEngineExcel:
    dataSheet = ""

    @staticmethod
    def ScanFileForPatterns(excelPathName):
        dataSheet = ExcelReader.readExcelFile(excelPathName)
        officeVersion = ""
        winVersion = ""
        for rowNumber in range(0, dataSheet.nrows - 1):
            for cell in dataSheet.row_slice(rowNumber):
                if str(cell.value).find('Office') != -1:
                    officeVersion = cell.value[str(cell.value).find("Office"):str(cell.value).find("Office") + 30]
                if str(cell.value).find('Win') != -1:
                    winVersion = cell.value[str(cell.value).find("Win"):str(cell.value).find("Win") + 30]
                if str(cell.value).find('W7') != -1:
                    winVersion = cell.value[str(cell.value).find('W7'):str(cell.value).find('W7') + 30]
                if str(cell.value).find('W8') != -1:
                    winVersion = cell.value[str(cell.value).find('W8'):str(cell.value).find('W8') + 30]
                if str(cell.value).find('W10') != -1:
                    winVersion = cell.value[str(cell.value).find('W10'):str(cell.value).find('W10') + 30]

        print("Office: ", officeVersion)
        print("Win: ", winVersion)
        return winVersion, officeVersion

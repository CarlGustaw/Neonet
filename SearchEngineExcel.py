import ExcelReader


class SearchEngineExcel:
    excelPathName = ""
    dataSheet = ""
    winDict = {
        "W7": "Windows 7",
        "W7P": "Windows 7 Pro",
        "W8": "Windows 8",
        "W10": "Windows 10",
        "W10P": "Windows 10 Pro"
    }

    def __init__(self, excelPathName):
        self.excelPathName = excelPathName
        self.dataSheet = ExcelReader.readExcelFile(self.excelPathName)
        print("Loaded excel files to search engine")

    def ScanFileForPatterns(self):
        officeVersion = ""
        winVersion = ""
        for rowNumber in range(0, self.dataSheet.nrows - 1):
            for cell in self.dataSheet.row_slice(rowNumber):
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
        return winVersion, officeVersion

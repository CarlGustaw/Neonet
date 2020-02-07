import ExcelReader


class SearchEngineExcel:
    excelPathName = ""
    dataSheet = ""
    officeVersion = ""
    winVersion = ""
    winDict = {
        "W7": "Windows 7",
        "W7P": "Windows 7 Pro",
        "W8": "Windows 8",
        "W8P": "Windows 8",
        "W10": "Windows 10",
        "W10P": "Windows 10 Pro"
    }

    def __init__(self, excelPathName):
        self.excelPathName = excelPathName
        self.dataSheet = ExcelReader.readExcelFile(self.excelPathName)
        print("Loaded excel files to search engine")

    def ScanFileForPatterns(self):

        for rowNumber in range(0, self.dataSheet.nrows - 1):
            for cell in self.dataSheet.row_slice(rowNumber):
                if str(cell.value).find('Office') != -1:
                    self.officeVersion = cell.value[str(cell.value).find("Office"):str(cell.value).find("Office") + 30]
                if str(cell.value).find('Windows') != -1 or str(cell.value).find('Win') != -1 or str(cell.value).find('W') != -1:
                    if str(cell.value).find('W7') != -1 or str(cell.value).find('Win7') != -1 or str(cell.value).find(
                            'Windows 7') != -1:
                        if str(cell.value).find('W7P') != -1 or str(cell.value).find('Pro') != -1:
                            self.winVersion = self.winDict.get("W7P")
                        else:
                            self.winVersion = self.winDict.get("W7")
                    if str(cell.value).find('W8') != -1 or str(cell.value).find('Windows 8') != -1:
                        if str(cell.value).find('W8P') != -1 or str(cell.value).find('Pro') != -1:
                            self.winVersion = self.winDict.get("W8P")
                        else:
                            self.winVersion = self.winDict.get("W8")
                    if str(cell.value).find('Windows 10') != -1 or str(cell.value).find('W10') != -1 or str(cell.value).find('Win10') != -1:
                        if str(cell.value).find('Windows 10 Pro') != -1 or str(cell.value).find('Pro') != -1 or str(cell.value).find("W10P") != -1:
                            self.winVersion = self.winDict.get("W10P")
                        else:
                            self.winVersion = self.winDict.get("W10")
        return self.winVersion, self.officeVersion

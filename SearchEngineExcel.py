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
    officeDict = {
        "O_2007": "Office 2007",
        "O_2010": "Office 2010",
        "O_2013": "Office 2013",
        "O_2016": "Office 2016"
    }

    ERRORMESSAGE = "Nie udalo sie odczytac wersj"

    def __init__(self, excelPathName):
        self.excelPathName = excelPathName
        self.dataSheet = ExcelReader.readExcelFile(self.excelPathName)
        print("Loaded excel files to search engine")

    def ScanFileForPatterns(self):

        for rowNumber in range(0, self.dataSheet.nrows - 1):
            for cell in self.dataSheet.row_slice(rowNumber):
                # Searching for office version
                if str(cell.value).find('Office') != -1:
                    if str(cell.value).find('2007') != -1:
                        self.officeVersion = self.officeDict.get("O_2007")
                    elif str(cell.value).find('2010') != -1:
                        self.officeVersion = self.officeDict.get("O_2010")
                    elif str(cell.value).find('2013') != -1:
                        self.officeVersion = self.officeDict.get("O_2013")
                    elif str(cell.value).find('2016') != -1:
                        self.officeVersion = self.officeDict.get("O_2016")
                    else:
                        self.officeVersion = cell.value[str(cell.value).find("Office"):str(cell.value).find("Office") + 30]

                # Searching for windows version
                if str(cell.value).find('Windows') != -1 or str(cell.value).find('Win') != -1 or str(cell.value).find(
                        'W') != -1:
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
                    if str(cell.value).find('Windows 10') != -1 or str(cell.value).find('W10') != -1 or str(
                            cell.value).find('Win10') != -1:
                        if str(cell.value).find('Windows 10 Pro') != -1 or str(cell.value).find('Pro') != -1 or str(
                                cell.value).find("W10P") != -1:
                            self.winVersion = self.winDict.get("W10P")
                        else:
                            self.winVersion = self.winDict.get("W10")

        # If no version was found error message is written
        if self.winVersion == "":
            self.winVersion = self.ERRORMESSAGE
        if self.officeVersion == "":
            self.officeVersion = self.ERRORMESSAGE

        return self.winVersion, self.officeVersion

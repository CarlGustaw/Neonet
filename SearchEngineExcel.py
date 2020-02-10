import ExcelReader


class SearchEngineExcel:
    excelPathName = ""
    dataSheet = ""
    officeVersion = ""
    numberOfOffices = 0
    winVersion = ""
    numberOfWindows = 0
    winDict = {
        "WXP": "Windows XP",
        "WXPP": "Windows XP Pro",
        "WV": "Windows Vista",
        "WVP": "Windows Vista Pro",
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

    patternsForWinXP = ["windows xp", "wxp", 'winxp', 'win xp', "xp"]
    patternsForWinVista = ["windows vista", "wv", "winvista", "win vista", "vis", "vista", "winv"]
    patternsForWin7 = ["windows 7", "w7", 'win7', 'win 7']
    patternsForWin8 = ["windows 8", "w8", 'win8', 'win 8']
    patternsForWin10 = ["windows 10", "w10", 'win10', 'win 10']

    patternsForIndexError = ["i", "I", "j", "J", "|", "L", "f", "F", "£"]

    ERRORMESSAGE = "Nie udalo sie odczytac wersji"

    def __init__(self, excelPathName):
        self.excelPathName = excelPathName
        self.dataSheet = ExcelReader.readExcelFile(self.excelPathName)
        print("Loaded excel files to search engine")

    def ScanFileForPatterns(self):
        print("Scan for patterns")
        for rowNumber in range(0, self.dataSheet.nrows - 1):
            for cell in self.dataSheet.row_slice(rowNumber):
                cellStringValue = str.lower(str(cell.value))
                # Searching for office version
                if cellStringValue.find('office') != -1:
                    if cellStringValue.find('szt') != -1:
                        valueOfszt = str(cellStringValue)[
                                     str.lower(str(cell.value)).find("szt") - 2: str.lower(str(cell.value)).find(
                                         "szt") + 3]
                        valueOfszt = list(valueOfszt)
                        try:
                            if valueOfszt[0] == "i" or valueOfszt[0] == "I" or valueOfszt[0] == "|" or valueOfszt[
                                0] == "j" or valueOfszt[0] == "f" or valueOfszt[
                                0] == "L":
                                valueOfszt[0] = 1
                                self.numberOfOffices = int(valueOfszt[0])
                                print("FOUND OFFICE \"szt\" in the same line How many of \"szt\":  ",
                                      self.numberOfOffices)
                        except:
                            print("IndexError: list index out of range", cellStringValue)

                    for minicell in self.dataSheet.row_slice(rowNumber - 2):
                        if str.lower(str(minicell.value)).find(" szt") != -1:
                            valueOfszt = str(minicell)[str(minicell).find("szt") - 2: str(minicell).find("szt") + 3]
                            valueOfszt = list(valueOfszt)
                            try:
                                if valueOfszt[0] == "i" or valueOfszt[0] == "I" or valueOfszt[0] == "|" or valueOfszt[
                                    0] == "j" or valueOfszt[0] == "f" or valueOfszt[
                                    0] == "L":
                                    valueOfszt[0] = 1
                                    self.numberOfOffices = int(valueOfszt[0])
                                    print("FOUND OFFICE \"szt\" in previous line  How many of \"szt\":  ",
                                          int(valueOfszt[0]))
                            except:
                                print("IndexError: list index out of range", cellStringValue)

                    if cellStringValue.find('2007') != -1:
                        self.officeVersion = self.officeDict.get("O_2007")
                    elif cellStringValue.find('2010') != -1:
                        self.officeVersion = self.officeDict.get("O_2010")
                    elif cellStringValue.find('2013') != -1:
                        self.officeVersion = self.officeDict.get("O_2013")
                    elif cellStringValue.find('2016') != -1:
                        self.officeVersion = self.officeDict.get("O_2016")
                    else:
                        self.officeVersion = cell.value[
                                             cellStringValue.find("office"):cellStringValue.find("office") + 30]

                # Searching for windows version
                if cellStringValue.find('windows') != -1 or cellStringValue.find('win') != -1 or cellStringValue.find(
                        'vis') != -1 or cellStringValue.find('vb') != -1:
                    if cellStringValue.find('szt') != -1:
                        valueOfszt = str(cellStringValue)[
                                     str.lower(str(cell.value)).find("szt") - 2: str.lower(str(cell.value)).find(
                                         "szt") + 3]
                        valueOfszt = list(valueOfszt)
                        try:
                            if valueOfszt[0] == "i" or valueOfszt[0] == "j" or valueOfszt[0] == "|" or valueOfszt[
                                0] == "£":
                                valueOfszt[0] = 1
                            self.numberOfWindows = int(valueOfszt[0])
                            print("FOUND WIN \"szt\" in the same line How many of \"szt\":  ", self.numberOfWindows)
                        except:
                            print("IndexError: list index out of range", cellStringValue)

                    self.searchForQuantityInTwoRowsHigher(self.dataSheet, rowNumber, cellStringValue)

                    self.searchForWindows(cellStringValue, self.patternsForWinXP, "WXPP", "WXP")
                    self.searchForWindows(cellStringValue, self.patternsForWinVista, "WVP", "WV")
                    self.searchForWindows(cellStringValue, self.patternsForWin7, "W7P", "W7")
                    self.searchForWindows(cellStringValue, self.patternsForWin8, "W8P", "W8")
                    self.searchForWindows(cellStringValue, self.patternsForWin10, "W10P", "W10")

        self.ifNoVersionFoundSetErrorMessage()
        self.showInformationFoundAboutWindowsAndOfficeVersion()
        return self.winVersion, self.officeVersion, self.numberOfOffices, self.numberOfWindows

    def searchForQuantityInTwoRowsHigher(self, dataSheet, rowNumber, cellStringValue):
        for earlierCells in self.dataSheet.row_slice(rowNumber - 2):
            if str.lower(str(earlierCells.value)).find(" szt") != -1:
                valueOfQuantity = self.setValueOfQuantity(earlierCells)
                valueOfQuantity = self.changeTypeOfValueOfQuantityToList(valueOfQuantity)
                self.ifIndexErrorOccursChangeItToOne(cellStringValue, self.patternsForIndexError, valueOfQuantity)

    def setValueOfQuantity(self, earlierCells):
        return str(earlierCells)[str(earlierCells).find("szt") - 2: str(earlierCells).find("szt") + 3]

    def changeTypeOfValueOfQuantityToList(self, valueOfQuantity):
        valueOfQuantity = list(valueOfQuantity)
        return valueOfQuantity

    def ifIndexErrorOccursChangeItToOne(self, cellStringValue, patternsForIndexError, valueOfQuantity):
        try:
            for pattern in patternsForIndexError:
                if pattern == valueOfQuantity[0]:
                    valueOfQuantity[0] = 1
                    self.numberOfWindows = int(valueOfQuantity[0])
                    print("FOUND WIN \"szt\" in previous line  How many of \"szt\":  ", self.numberOfWindows)
        except IndexError:
            print("List index out of range, showing whole cell value:   ", cellStringValue)

    # Method take as cell value as argument, specify pattern list to search and two dictionary links to version type.
    def searchForWindows(self, cellStringValue, patternList, dictLinkIfPro, dictLinkIfNotPro):
        for pattern in patternList:
            if cellStringValue.find(pattern):
                self.searchIfVersionIsProfessional(cellStringValue, dictLinkIfPro, dictLinkIfNotPro)

    def searchIfVersionIsProfessional(self, cellStringValue, dictLinkIfPro, dictLinkIfNotPro):
        if cellStringValue.find('pro') != -1:
            self.winVersion = self.winDict.get(dictLinkIfPro)
        else:
            self.winVersion = self.winDict.get(dictLinkIfNotPro)

    def ifNoVersionFoundSetErrorMessage(self):
        if self.winVersion == "":
            self.winVersion = self.ERRORMESSAGE
        if self.officeVersion == "":
            self.officeVersion = self.ERRORMESSAGE

    def showInformationFoundAboutWindowsAndOfficeVersion(self):
        print("From SearchEngine: Office version: ", self.officeVersion, "   Windows Version:  ", self.winVersion)

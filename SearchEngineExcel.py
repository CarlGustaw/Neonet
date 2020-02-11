import ExcelReader


class SearchEngineExcel:
    patternWindowsFound = False
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
        "W10": "Windows 10 Home",
        "W10P": "Windows 10 Pro",
        "Weird": "Weird"
    }
    officeDict = {
        "2007": "Office 2007",
        "2010": "Office 2010",
        "2013": "Office 2013",
        "2016": "Office 2016"
    }

    patternsWeirdAndUncommon = ["xp prg+win"]
    patternsForOffice = ["2007", "2010", "2013", "2016"]
    patternsForWinXP = ["windows xp", "wxp", 'winxp', 'win xp', "xp"]
    patternsForWinVista = ["windows vista", "wv", "winvista", "win vista", "vis", "vista", "winv", "vistabusiness"]
    patternsForWin7 = ["windows 7", "w7", 'win7', 'win 7']
    patternsForWin8 = ["windows 8", "w8", 'win8', 'win 8']
    patternsForWin10 = ["windows 10", "w10", "win10", "win 10", "win1"]

    patternsForIndexError = ["i", "I", "j", "J", "|", "L", "f", "F", "£"]

    ERROR_MESSAGE = "Nie udalo sie odczytac wersji"

    def __init__(self, excelPathName):
        self.excelPathName = excelPathName
        self.dataSheet = ExcelReader.readExcelFile(self.excelPathName)
        print("Loaded excel files to search engine")

    def ScanFileForPatterns(self):
        print("Scan for patterns")
        rowOfficeValue = []
        rowWindowsValue = []
        for rowNumber in range(0, self.dataSheet.nrows - 1):
            for cell in self.dataSheet.row_slice(rowNumber):
                cellStringValue = str.lower(str(cell.value))

                # Searching for office version
                if cellStringValue.find('office') != -1:
                    rowOfficeValue.append(self.dataSheet.row_slice(rowNumber))
                    print("ROW VALUE OFFICE:::: ", rowOfficeValue)
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

                    # Searching for office version
                    self.searchForPatternIn(cellStringValue, self.patternsForOffice, "", "")

                # Searching  for windows version
                self.searchEngineForWindows(cellStringValue, rowNumber, rowWindowsValue)

        self.ifNoVersionFoundSetErrorMessage()
        self.showInformationFoundAboutWindowsAndOfficeVersion()
        return self.winVersion, self.officeVersion, self.numberOfOffices, self.numberOfWindows

    def searchEngineForWindows(self, cellStringValue, currentRowNumber, rowWindowsValue):

        # Searching if in line is any trace of word windows or vista (some excels don't read pdfs correctly)
        if cellStringValue.find('windows') != -1 or cellStringValue.find('win') != -1 or cellStringValue.find(
                'vis') != -1 or cellStringValue.find('vb') != -1:

            rowWindowsValue.append(self.dataSheet.row_slice(currentRowNumber))
            print("ROW VALUE WINDOWS:::: ", rowWindowsValue)

            self.searchForQuantityInTwoRowsHigherOrSameLine("SameLine", currentRowNumber, cellStringValue, "Win")
            self.searchForQuantityInTwoRowsHigherOrSameLine("NotTheSameLine", currentRowNumber, cellStringValue, "Win")
            print("Windows found state:  ", self.patternWindowsFound)
            if self.patternWindowsFound == False:
                self.searchForPatternIn(cellStringValue, self.patternsWeirdAndUncommon, "Weird", "")
            if self.patternWindowsFound == False:
                self.searchForPatternIn(cellStringValue, self.patternsForWinXP, "WXPP", "WXP")
            if self.patternWindowsFound == False:
                self.searchForPatternIn(cellStringValue, self.patternsForWinVista, "WVP", "WV")
            if self.patternWindowsFound == False:
                self.searchForPatternIn(cellStringValue, self.patternsForWin7, "W7P", "W7")
            if self.patternWindowsFound == False:
                self.searchForPatternIn(cellStringValue, self.patternsForWin8, "W8P", "W8")
            if self.patternWindowsFound == False:
                self.searchForPatternIn(cellStringValue, self.patternsForWin10, "W10P", "W10")

    def searchForQuantityInTwoRowsHigherOrSameLine(self, setWhichLine, currentRowNumber, cellStringValue, WinOrOffice):
        # Search for quantity in two lines above current row
        if setWhichLine != "SameLine":
            for cell in self.dataSheet.row_slice(currentRowNumber - 2):
                cellsLowerStringValue = str.lower(str(cell.value))
                self.searchForQuantityMarkAndCheckingForAnyErrorInIndex(cellsLowerStringValue, WinOrOffice)
        else:
            # Search for quantity in the same line
            self.searchForQuantityMarkAndCheckingForAnyErrorInIndex(cellStringValue, WinOrOffice)

    # If mark for quantity was found, is gonna be write down as list of characters
    def searchForQuantityMarkAndCheckingForAnyErrorInIndex(self, searchedCell, WinOrOffice):
        if searchedCell.find('szt') != -1:
            valueOfQuantity = self.setValueOfQuantity(searchedCell)
            valueOfQuantity = self.changeTypeOfValueOfQuantityToList(valueOfQuantity)
            # Any wrongly read quantity is gonna fixed, by switching first index to one
            self.ifIndexErrorOccursChangeItToOneAndSetQuantities(searchedCell, self.patternsForIndexError, valueOfQuantity, WinOrOffice)

    # Return the reading frame with quantity number in it
    def setValueOfQuantity(self, cell):
        return str(cell)[str(cell).find("szt") - 2: str(cell).find("szt") + 3]

    def changeTypeOfValueOfQuantityToList(self, valueOfQuantity):
        valueOfQuantity = list(valueOfQuantity)
        return valueOfQuantity

    def ifIndexErrorOccursChangeItToOneAndSetQuantities(self, cellStringValue, patternsForIndexError, valueOfQuantity,
                                                        WinOrOffice):
        try:
            for pattern in patternsForIndexError:
                if pattern == valueOfQuantity[0]:
                    valueOfQuantity[0] = 1
                    self.setQuantitiesForWindowsOrOffice(WinOrOffice, valueOfQuantity)
        except IndexError:
            print("List index out of range, showing whole cell value:   ", cellStringValue)

    def setQuantitiesForWindowsOrOffice(self, WinOrOffice, valueOfQuantity):
        if WinOrOffice == "Win":
            self.numberOfWindows = int(valueOfQuantity[0])
            print("Found Windows \"szt\". How many of \"szt\":  ", self.numberOfWindows)
        else:
            self.numberOfOffices = int(valueOfQuantity[0])
            print("Found Office \"szt\". How many of \"szt\":  ", self.numberOfOffices)

    # Method take as cell value as argument, specify pattern list to search and two dictionary links to version type.
    def searchForPatternIn(self, cellStringValue, patternList, dictLinkIfPro, dictLinkIfNotPro):

        if dictLinkIfPro == "Weird" and dictLinkIfNotPro == "":
            for pattern in patternList:
                if cellStringValue.find(pattern) != -1:
                    self.patternWindowsFound = True
                    self.winVersion = pattern
        # If no dictionary link was provided, turn on module for search Office version
        elif dictLinkIfPro == "" and dictLinkIfNotPro == "":
            for pattern in patternList:
                if cellStringValue.find(pattern) != -1:
                    self.officeVersion = self.officeDict.get(pattern)
                else:
                    self.officeVersion = cellStringValue
        else:
            # If dictionary link was provided, turn on module for search Windows version
            for pattern in patternList:
                if cellStringValue.find(pattern) != -1:
                    self.patternWindowsFound = True
                    print("Wartość:  ", cellStringValue)
                    self.searchIfVersionIsProfessionalOrNot(cellStringValue, dictLinkIfPro, dictLinkIfNotPro)
                    print("Pattern found :  ", pattern)
                    break
                else:
                    self.winVersion = cellStringValue

    def searchIfVersionIsProfessionalOrNot(self, cellStringValue, dictLinkIfPro, dictLinkIfNotPro):
        if cellStringValue.find('pro') != -1 or cellStringValue.find('p') != -1 and cellStringValue.find('prod') == -1:
            self.winVersion = self.winDict.get(dictLinkIfPro)
        else:
            self.winVersion = self.winDict.get(dictLinkIfNotPro)

    def ifNoVersionFoundSetErrorMessage(self):
        if self.winVersion == "":
            self.winVersion = self.ERROR_MESSAGE
        if self.officeVersion == "":
            self.officeVersion = self.ERROR_MESSAGE

    def showInformationFoundAboutWindowsAndOfficeVersion(self):
        print("From SearchEngine: Office version: ", self.officeVersion, "   Windows Version:  ", self.winVersion)

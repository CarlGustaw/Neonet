import ExcelReader


class SearchEngineExcel:
    patternWindowsFound = False
    excelPathName = ""
    dataSheet = ""
    officeVersion = ""
    winVersion = ""
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

    patternsForWindows = ["windows", "win", "vis", "vb", "wlO", "w7", "w8", "w10", "wxp"]
    patternsToAvoid = ["wr", "wy", "wys", "wa", "wie", "wn", "wh", "wa"]
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

    def ScanFileForPatterns(self):
        listRowsOfficeValue = []
        listRowWindowsValue = []
        for rowNumber in range(0, self.dataSheet.nrows - 1):
            for cell in self.dataSheet.row_slice(rowNumber):
                cellStringValue = str.lower(str(cell.value))

                # Searching for office version and writing down whole rows
                if cellStringValue.find('off') != -1 and cellStringValue.find(
                        "officejet") == -1 and cellStringValue.find("officepower") == -1 or cellStringValue.find(
                        'otllce') != -1:
                    listRowsOfficeValue.append(self.dataSheet.row_slice(rowNumber))

                    self.searchForPatternIn(cellStringValue, self.patternsForOffice, "", "")

                # Searching for windows version and writing down whole rows
                for pattern in self.patternsForWindows:
                    if cellStringValue.find(pattern) != -1:
                        listRowWindowsValue.append(self.dataSheet.row_slice(rowNumber))

                # Searching  for windows version
                listRowWindowsValue.append(self.searchEngineForWindows(cellStringValue, rowNumber))

        self.ifNoVersionFoundSetErrorMessage()
        self.showInformationFoundAboutWindowsAndOfficeVersion()
        return self.winVersion, self.officeVersion, listRowWindowsValue, listRowsOfficeValue

    def searchEngineForWindows(self, cellStringValue, currentRowNumber):

        # Searching if in line is any trace of word windows or vista (some excels don't read pdfs correctly)
        if cellStringValue.find('windows') != -1 or cellStringValue.find('win') != -1 or cellStringValue.find(
                'vis') != -1 or cellStringValue.find('vb') != -1 or cellStringValue.find("WlO") != -1:

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

    # Method take as cell value as argument, specify pattern list to search and two dictionary links to version type.
    def searchForPatternIn(self, cellStringValue, patternList, dictLinkIfPro, dictLinkIfNotPro):
        # If no dictionary link was provided, turn on module for search Office version
        if dictLinkIfPro == "" and dictLinkIfNotPro == "":
            for pattern in patternList:
                if cellStringValue.find(pattern) != -1:
                    self.officeVersion = self.officeDict.get(pattern)
                    break
                else:
                    self.officeVersion = cellStringValue
        else:
            # If dictionary link was provided, turn on module for search Windows version
            for pattern in patternList:
                if cellStringValue.find(pattern) != -1:
                    self.patternWindowsFound = True
                    self.searchIfVersionIsProfessionalOrNot(cellStringValue, dictLinkIfPro, dictLinkIfNotPro)
                    break

    def searchIfVersionIsProfessionalOrNot(self, cellStringValue, dictLinkIfPro, dictLinkIfNotPro):
        if cellStringValue.find('pro') != -1 and cellStringValue.find('prod') == -1:
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

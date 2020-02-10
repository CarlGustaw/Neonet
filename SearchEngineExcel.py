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
                        valueOfszt = str(cellStringValue)[str.lower(str(cell.value)).find("szt") - 2: str.lower(str(cell.value)).find("szt") + 3]
                        valueOfszt = list(valueOfszt)
                        if valueOfszt[0] == "i":
                            valueOfszt[0] = 1
                        print("FOUND \"szt\" in the same line How many of \"szt\":  ", valueOfszt[0])

                    for minicell in self.dataSheet.row_slice(rowNumber - 2):
                        if str.lower(str(minicell.value)).find(" szt") != -1:
                            valueOfszt = str(minicell)[str(minicell).find("szt") - 2: str(minicell).find("szt") + 3]
                            valueOfszt = list(valueOfszt)
                            print("FOUND \"szt\" in previous line  How many of \"szt\":  ", valueOfszt[0])
                    if cellStringValue.find('2007') != -1:
                        self.officeVersion = self.officeDict.get("O_2007")
                        break
                    elif cellStringValue.find('2010') != -1:
                        self.officeVersion = self.officeDict.get("O_2010")
                        break
                    elif cellStringValue.find('2013') != -1:
                        self.officeVersion = self.officeDict.get("O_2013")
                        break
                    elif cellStringValue.find('2016') != -1:
                        self.officeVersion = self.officeDict.get("O_2016")
                        break
                    else:
                        self.officeVersion = cell.value[
                                             cellStringValue.find("office"):cellStringValue.find("office") + 30]

                # Searching for windows version
                if cellStringValue.find('windows') != -1 or cellStringValue.find('win') != -1 or cellStringValue.find(
                        'w') != -1 or cellStringValue.find('vis') != -1:
                    if cellStringValue.find('wxp') != -1 or cellStringValue.find('winxp') != -1 or cellStringValue.find(
                            'windows xp') != -1 or str(cell.value).find('win xp') != -1:
                        if cellStringValue.find('pro') != -1:
                            self.winVersion = self.winDict.get("WXPP")
                            break
                        else:
                            self.winVersion = self.winDict.get("WXP")
                            break
                    if cellStringValue.find('wv') != -1 or cellStringValue.find('winv') != -1 or cellStringValue.find(
                            'windows vista') != -1 or cellStringValue.find('win v') != -1 or cellStringValue.find(
                        'vis') != -1:
                        if cellStringValue.find('Pro') != -1:
                            self.winVersion = self.winDict.get("WVP")
                            break
                        else:
                            self.winVersion = self.winDict.get("WV")
                            break
                    if cellStringValue.find('w7') != -1 or cellStringValue.find('win7') != -1 or cellStringValue.find(
                            'windows 7') != -1 or str(cell.value).find('win 7') != -1:
                        if cellStringValue.find('w7p') != -1 or cellStringValue.find('pro') != -1:
                            self.winVersion = self.winDict.get("W7P")
                            break
                        else:
                            self.winVersion = self.winDict.get("W7")
                            break
                    if cellStringValue.find('w8') != -1 or cellStringValue.find('windows 8') != -1 or str(
                            cell.value).find('win 8') != -1 or cellStringValue.find('win8') != -1:
                        if cellStringValue.find('w8P') != -1 or cellStringValue.find('pro') != -1:
                            self.winVersion = self.winDict.get("w8p")
                            break
                        else:
                            self.winVersion = self.winDict.get("w8")
                            break
                    if cellStringValue.find('windows 10') != -1 or cellStringValue.find('w10') != -1 or str(
                            cell.value).find('win10') != -1 or cellStringValue.find('win 10') != -1:
                        if cellStringValue.find('windows 10 pro') != -1 or cellStringValue.find('pro') != -1 or str(
                                cell.value).find("w10P") != -1:
                            self.winVersion = self.winDict.get("W10P")
                            break
                        else:
                            self.winVersion = self.winDict.get("W10")
                            break

        # If no version was found error message is written
        if self.winVersion == "":
            self.winVersion = self.ERRORMESSAGE
        if self.officeVersion == "":
            self.officeVersion = self.ERRORMESSAGE
        print("From SearchEngine: Office version: ", self.officeVersion, "   Windows Version:  ", self.winVersion)
        return self.winVersion, self.officeVersion

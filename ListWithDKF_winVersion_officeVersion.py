from SearchEngineExcel import SearchEngineExcel


class ListWithDKF_winVersion_officeVersion:
    ListDKF_WIN_OFFICE = []

    def __init__(self):
        print("Creating ExcelReader")

    def addFoundPattern(self, id_DKF, excelPath):
        winVersion, officeVersion, numberOfOffices, numberOfWindows, listRowsWindows, listRowsOffice = self.getPatternsFromSearchEngine(excelPath)
        listRowsWindows = [i for i in listRowsWindows if i]
        self.ListDKF_WIN_OFFICE.insert(len(self.ListDKF_WIN_OFFICE),
                                       [id_DKF, winVersion, officeVersion, numberOfOffices, numberOfWindows, listRowsWindows, listRowsOffice])

    @staticmethod
    def getPatternsFromSearchEngine(excelPath):
        SearchEngine = SearchEngineExcel(excelPath)
        winVersion, officeVersion, numberOfOffices, numberOfWindows, listRowsWindows, listRowsOffice = SearchEngine.ScanFileForPatterns()
        return winVersion, officeVersion, numberOfOffices, numberOfWindows, listRowsWindows, listRowsOffice

    def getActualList(self):
        return self.ListDKF_WIN_OFFICE

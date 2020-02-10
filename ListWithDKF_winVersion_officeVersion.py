from SearchEngineExcel import SearchEngineExcel


class ListWithDKF_winVersion_officeVersion:
    ListDKF_WIN_OFFICE = []

    def __init__(self):
        print("Creating ExcelReader")

    def addFoundPattern(self, id_DKF, excelPath):
        winVersion, officeVersion, numberOfOffices, numberOfWindows = self.getPatternsFromSearchEngine(excelPath)
        self.ListDKF_WIN_OFFICE.insert(len(self.ListDKF_WIN_OFFICE),
                                       [id_DKF, winVersion, officeVersion, numberOfOffices, numberOfWindows])

    @staticmethod
    def getPatternsFromSearchEngine(excelPath):
        SearchEngine = SearchEngineExcel(excelPath)
        winVersion, officeVersion, numberOfOffices, numberOfWindows = SearchEngine.ScanFileForPatterns()
        return winVersion, officeVersion, numberOfOffices, numberOfWindows

    def getActualList(self):
        return self.ListDKF_WIN_OFFICE

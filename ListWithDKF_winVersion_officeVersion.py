from SearchEngineExcel import SearchEngineExcel


class ListWithDKF_winVersion_officeVersion:
    ListDKF_WIN_OFFICE = []

    def __init__(self):
        print("Creating ExcelReader")

    def addFoundPattern(self, id_DKF, excelPath):
        winVersion, officeVersion = self.getPatternsFromSearchEngine(excelPath)
        self.ListDKF_WIN_OFFICE.insert(len(self.ListDKF_WIN_OFFICE), [id_DKF, winVersion, officeVersion])
        print("Actual list: ", self.ListDKF_WIN_OFFICE)
        print()

    @staticmethod
    def getPatternsFromSearchEngine(excelPath):
        SearchEngine = SearchEngineExcel(excelPath)
        winVersion, officeVersion = SearchEngine.ScanFileForPatterns()
        return winVersion, officeVersion

    def getActualList(self):
        return self.ListDKF_WIN_OFFICE
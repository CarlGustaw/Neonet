from SearchEngineExcel import SearchEngineExcel
from ExcelPathsGetter import FilesInDir


class ListWithDKF_winVersion_officeVersion:
    excelPathsGetter = ""
    ListDKF_WIN_OFFICE = []

    def __init__(self, DIRWITHPDFCHANGEDTOEXCEL):
        print("Creating ExcelReader")
        self.excelPathsGetter = FilesInDir(DIRWITHPDFCHANGEDTOEXCEL)

    def makeList(self, id_DKF):
        print("Given id_DKF: ", id_DKF)
        for excelPath in self.excelPathsGetter.getFilesPaths():
            if excelPath.find(str(id_DKF)) != -1:
                print("Correct path found   ", id_DKF, "    ", excelPath)
                SearchEngine = SearchEngineExcel(excelPath)
                winVersion, officeVersion = SearchEngine.ScanFileForPatterns()
                self.ListDKF_WIN_OFFICE.insert(len(self.ListDKF_WIN_OFFICE), [id_DKF, winVersion, officeVersion])
        print("Actual list: ", self.ListDKF_WIN_OFFICE)
        return self.ListDKF_WIN_OFFICE

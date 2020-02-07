from SearchEngineExcel import SearchEngineExcel
from ExcelPathsGetter import FilesInDir


class ListWithDKF_winVersion_officeVersion:
    excelPathsGetter = ""
    ListDKF_WIN_OFFICE = []

    def __init__(self, DIRWITHPDFCHANGEDTOEXCEL):
        print("Creating ExcelReader")
        self.excelPathsGetter = FilesInDir(DIRWITHPDFCHANGEDTOEXCEL)
        print("Number of files in dir:  ", len(self.excelPathsGetter.getFilesPaths()))

    def makeList(self, id_DKF):
        for excelPath in self.excelPathsGetter.getFilesPaths():
            if excelPath.find(str(id_DKF)) != -1:
                print("Correct path found   ", id_DKF, "    ", excelPath)
                SearchEngine = SearchEngineExcel(excelPath)
                winVersion, officeVersion = SearchEngine.ScanFileForPatterns()
                print("Wersja Win:  ", winVersion, "    Wersja Office:      ", officeVersion)
                self.ListDKF_WIN_OFFICE.insert(len(self.ListDKF_WIN_OFFICE), [id_DKF, winVersion, officeVersion])
                print("Actual list: ", self.ListDKF_WIN_OFFICE)
        return self.ListDKF_WIN_OFFICE

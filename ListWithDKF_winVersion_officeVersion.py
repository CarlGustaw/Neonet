from RowToObjects import RowMaker
from SearchEngineExcel import SearchEngineExcel
from ExcelPathsGetter import FilesInDir


class ListWithDKF_winVersion_officeVersion:
    MAINEXCELPATHNAME = ""
    excelPathsGetter = ""
    ListDKF_WIN_OFFICE = []

    def __init__(self, MAINEXCELPATHNAME, DIRWITHPDFCHANGEDTOEXCEL):
        self.MAINEXCELPATHNAME = MAINEXCELPATHNAME
        print("Creating ExcelReader")
        self.excelPathsGetter = FilesInDir(DIRWITHPDFCHANGEDTOEXCEL)

    def makeList(self, id_DKF):
        rowObjects, rowsWithBadDKF = RowMaker.readExcelFileToSheetAndMakingObject(self.MAINEXCELPATHNAME)
        print("Given id_DKF: ", id_DKF)

        for excelPath in self.excelPathsGetter.getFilesPaths():
            if excelPath.find(id_DKF) != -1:
                print("Correct path found   ", id_DKF, "    ", excelPath)
                SearchEngine = SearchEngineExcel(excelPath)
                winVersion, officeVersion = SearchEngine.ScanFileForPatterns()
                self.ListDKF_WIN_OFFICE.insert(len(self.ListDKF_WIN_OFFICE), [id_DKF, winVersion, officeVersion])

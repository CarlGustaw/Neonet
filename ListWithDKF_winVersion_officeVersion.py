from RowToObjects import RowMaker
from SearchEngineExcel import SearchEngineExcel
import ExcelPathsGetter


class ListWithDKF_winVersion_officeVersion:
    MAINEXCELPATHNAME = ""
    ListDKF_WIN_OFFICE = []

    def __init__(self, MAINEXCELPATHNAME):
        self.MAINEXCELPATHNAME = MAINEXCELPATHNAME

    def makeList(self, id_DKF):
        rowObjects, rowsWithBadDKF = RowMaker.readExcelFileToSheetAndMakingObject(self.MAINEXCELPATHNAME)
        print("Given id_DKF: ", id_DKF)
        for excelPath in ExcelPathsGetter.excelPaths:
            if excelPath.find(id_DKF) != -1:
                print("Correct path found   ", id_DKF, "    ", excelPath)
                winVersion, officeVersion = SearchEngineExcel.ScanFileForPatterns(excelPath)
                self.ListDKF_WIN_OFFICE.insert(len(self.ListDKF_WIN_OFFICE), [id_DKF, winVersion, officeVersion])

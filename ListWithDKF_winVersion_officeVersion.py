from RowToObjects import RowMaker
from SearchEngineExcel import SearchEngineExcel
import ExcelPathsGetter


class ListWithDKF_winVersion_officeVersion:
    MAINEXCELPATHNAME = ""
    ListDKF_WIN_OFFICE = []

    def __init__(self, MAINEXCELPATHNAME):
        self.MAINEXCELPATHNAME = MAINEXCELPATHNAME

    # Read DKFs from main excel file
    def __ReadDKFs(self):
        rowObjects, rowsWithBadDKF = RowMaker.readExcelFileToSheetAndMakingObject()
        return rowObjects, rowsWithBadDKF

    # make data format that excel writer will accept, contain DKS, windows version and office version
    def makeList(self, id_DKF):
        self.__ReadDKFs()
        print("Given id_DKF: ", id_DKF)
        print("Detect paths: ", ExcelPathsGetter.excelPaths)
        for excelPath in ExcelPathsGetter.excelPaths:
            if excelPath.find(id_DKF) != -1:
                print("Correct path found   ", id_DKF, "    ", excelPath)
                winVersion, officeVersion = SearchEngineExcel.ScanFileForPatterns(excelPath)
                self.ListDKF_WIN_OFFICE.insert(len(self.ListDKF_WIN_OFFICE), [id_DKF, winVersion, officeVersion])
        return ListWithDKF_winVersion_officeVersion

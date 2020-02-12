from RowToObjects import RowMaker
from ExcelPathsGetter import FilesInDir
from ListWithDKF_winVersion_officeVersion import ListWithDKF_winVersion_officeVersion


class BuildFinalExcel:
    MAINEXCELPATHNAME = ""
    index_of_dkf_column = 0
    index_row_scan = 0
    DIRWITHPDFCHANGEDTOEXCEL = ""
    FinalListMaker = ""
    finalListDKF_WIN_OFFICE = []
    uniqueList = []

    def __init__(self, MAINEXCELPATHNAME, index_of_dkf_column, index_row_scan, DIRWITHPDFCHANGEDTOEXCEL):
        self.MAINEXCELPATHNAME = MAINEXCELPATHNAME
        self.index_of_dkf_column = index_of_dkf_column
        self.index_row_scan = index_row_scan
        self.DIRWITHPDFCHANGEDTOEXCEL = DIRWITHPDFCHANGEDTOEXCEL
        self.FinalListMaker = ListWithDKF_winVersion_officeVersion()
        print("Builder finished")

    def build(self):
        print("Building final excel")
        rowObjects, rowsWithBadDKF = RowMaker.readExcelFileToSheetAndMakingObject(self.MAINEXCELPATHNAME,
                                                                                  self.index_of_dkf_column,
                                                                                  self.index_row_scan)

        excelPaths = FilesInDir(self.DIRWITHPDFCHANGEDTOEXCEL).getFilesPaths()
        for excelPath in excelPaths:
            for correctDKF in rowObjects:
                if excelPath.find(correctDKF.getID_DKF()) != -1:
                    print()
                    print("Correct path found   ", correctDKF.getID_DKF(), "    ", excelPath)
                    self.FinalListMaker.addFoundPattern(correctDKF.getID_DKF(), excelPath)
        self.finalListDKF_WIN_OFFICE = self.FinalListMaker.getActualList()

    def showFinalList(self):
        print()
        print("Final List:  ", self.finalListDKF_WIN_OFFICE)

    def getFinalUniqueList(self):
        return self.finalListDKF_WIN_OFFICE

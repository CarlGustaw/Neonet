from RowToObjects import RowMaker
from ExcelPathsGetter import FilesInDir
from ListWithDKF_winVersion_officeVersion import ListWithDKF_winVersion_officeVersion


class BuildFinalExcel:
    MAINEXCELPATHNAME = ""
    DIRWITHPDFCHANGEDTOEXCEL = ""
    FinalListMaker = ""
    finalListDKF_WIN_OFFICE = []
    uniqueList = []

    def __init__(self, MAINEXCELPATHNAME, DIRWITHPDFCHANGEDTOEXCEL):
        self.MAINEXCELPATHNAME = MAINEXCELPATHNAME
        self.DIRWITHPDFCHANGEDTOEXCEL = DIRWITHPDFCHANGEDTOEXCEL
        self.FinalListMaker = ListWithDKF_winVersion_officeVersion()
        print("Builder finished")

    def build(self):
        print("Building final excel")
        rowObjects, rowsWithBadDKF = RowMaker.readExcelFileToSheetAndMakingObject(self.MAINEXCELPATHNAME)

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

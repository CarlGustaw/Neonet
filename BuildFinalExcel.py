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
        print("Builder finished")

    def build(self):
        print("Building final excel")
        rowObjects, rowsWithBadDKF = RowMaker.readExcelFileToSheetAndMakingObject(self.MAINEXCELPATHNAME)

        excelPaths = FilesInDir(self.DIRWITHPDFCHANGEDTOEXCEL).getFilesPaths()

        for excelPath in excelPaths:
            for correctDKF in rowObjects:
                if excelPath.find("90409") != -1 and correctDKF.getID_DKF() == "90409":
                    print("Dkf: ", correctDKF.getID_DKF(), " path:", excelPath)
        # self.FinalListMaker = ListWithDKF_winVersion_officeVersion(self.DIRWITHPDFCHANGEDTOEXCEL)
        # for correctDKF in rowObjects:
        # self.finalListDKF_WIN_OFFICE.append(self.FinalListMaker.makeList(correctDKF.getID_DKF()))

    def showFinalList(self):
        flattened = [val for sublist in self.finalListDKF_WIN_OFFICE for val in sublist]
        print("Final List:  ", flattened)

    def getFinalUniqueList(self):
        flattened = [val for sublist in self.finalListDKF_WIN_OFFICE for val in sublist]
        return flattened

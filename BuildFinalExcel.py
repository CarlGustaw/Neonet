from RowToObjects import RowMaker
from ListWithDKF_winVersion_officeVersion import ListWithDKF_winVersion_officeVersion


class BuildFinalExcel:
    MAINEXCELPATHNAME = ""
    DIRWITHPDFCHANGEDTOEXCEL = ""
    FinalListMaker = ""
    finalListDKF_WIN_OFFICE = []

    def __init__(self, MAINEXCELPATHNAME, DIRWITHPDFCHANGEDTOEXCEL):
        self.MAINEXCELPATHNAME = MAINEXCELPATHNAME
        self.DIRWITHPDFCHANGEDTOEXCEL = DIRWITHPDFCHANGEDTOEXCEL
        print("Builder finished")

    def build(self):
        print("Building final excel")
        rowObjects, rowsWithBadDKF = RowMaker.readExcelFileToSheetAndMakingObject(self.MAINEXCELPATHNAME)
        self.FinalListMaker = ListWithDKF_winVersion_officeVersion(self.DIRWITHPDFCHANGEDTOEXCEL)

        for correctDKF in rowObjects:
            self.finalListDKF_WIN_OFFICE.append(self.FinalListMaker.makeList(correctDKF.getID_DKF()))

    def showFinalList(self):
        print("Final List:  ", self.finalListDKF_WIN_OFFICE)


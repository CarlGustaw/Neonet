from RowToObjects import RowMaker
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
        self.FinalListMaker = ListWithDKF_winVersion_officeVersion(self.DIRWITHPDFCHANGEDTOEXCEL)

        for correctDKF in rowObjects:
            self.finalListDKF_WIN_OFFICE.append(self.FinalListMaker.makeList(correctDKF.getID_DKF()))

        self.uniqueList = []
        for row in self.finalListDKF_WIN_OFFICE:
            if row not in self.uniqueList:
                self.uniqueList.append(row)

    def showFinalList(self):
        flattened = [val for sublist in self.uniqueList for val in sublist]
        print("Final List:  ",  flattened)

    def getFinalUniqueList(self):
        flattened = [val for sublist in self.uniqueList for val in sublist]
        return flattened


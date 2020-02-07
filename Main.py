from RowToObjects import RowMaker
from ListWithDKF_winVersion_officeVersion import ListWithDKF_winVersion_officeVersion
from Xlsxwriter import ExcelWriter

MAINEXCELPATHNAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/DoTestow.xlsx"
EXCELPATHNAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/90408.xls"
DIRWITHPDFCHANGEDTOEXCEL = "//kmsrv01/OCR/EXCEL/OUTPUT/*.xls"

rowObjects, rowsWithBadDKF = RowMaker.readExcelFileToSheetAndMakingObject(MAINEXCELPATHNAME)
rowObjects[0].showRowObject()
rowObjects[6].showRowObject()
print("Number of rows with bad DKF:  ", len(rowsWithBadDKF))
print()

finalList = ListWithDKF_winVersion_officeVersion(MAINEXCELPATHNAME, DIRWITHPDFCHANGEDTOEXCEL)
finalList.makeList("90408")
print("Row from final list: ", finalList.ListDKF_WIN_OFFICE[0])
print("DKF:  ", finalList.ListDKF_WIN_OFFICE[0][0])
print("Win:  ", finalList.ListDKF_WIN_OFFICE[0][1])
print("Office:  ", finalList.ListDKF_WIN_OFFICE[0][2])
ToExcel = ExcelWriter()
ToExcel.makeExcel(finalList.ListDKF_WIN_OFFICE[0])


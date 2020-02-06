from ListWithDKF_winVersion_officeVersion import ListWithDKF_winVersion_officeVersion
from Xlsxwriter import ExcelWriter

MAINEXCELPATHNAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/DoTestow.xlsx"
EXCELPATHNAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/90408.xls"

dataForExcelWriter = ListWithDKF_winVersion_officeVersion(MAINEXCELPATHNAME)
print("Row from dataForExcelWriter: ", dataForExcelWriter.ListDKF_WIN_OFFICE[0])
print("DKF:  ", dataForExcelWriter.ListDKF_WIN_OFFICE[0][0])
print("Win:  ", dataForExcelWriter.ListDKF_WIN_OFFICE[0][1])
print("Office:  ", dataForExcelWriter.ListDKF_WIN_OFFICE[0][2])

# Write data as excel file in dir (TestyPyKsiegowosc)
ToExcel = ExcelWriter()
ToExcel.makeExcel(dataForExcelWriter.makeList("90408"))

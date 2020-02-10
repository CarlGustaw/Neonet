from BuildFinalExcel import BuildFinalExcel
from Xlsxwriter import ExcelWriter

MAINEXCELPATHNAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/DoTestow.xlsx"
EXCELPATHNAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/90408.xls"
DIRWITHPDFCHANGEDTOEXCEL = "//kmsrv01/OCR/EXCEL/OUTPUT/*.xls"

builder = BuildFinalExcel(MAINEXCELPATHNAME, DIRWITHPDFCHANGEDTOEXCEL)
builder.build()
builder.showFinalList()

ToExcel = ExcelWriter()
ToExcel.makeExcel(builder.getFinalUniqueList())


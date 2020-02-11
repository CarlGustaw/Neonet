from BuildFinalExcel import BuildFinalExcel
from Xlsxwriter import ExcelWriter

MAINEXCELPATHNAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/DoTestow.xlsx"
DIRWITHPDFCHANGEDTOEXCEL = "C:/Users/dMichalczak/TestyPyKsiegowosc/Pdf_to_Excel/*.xls"

builder = BuildFinalExcel(MAINEXCELPATHNAME, DIRWITHPDFCHANGEDTOEXCEL)
builder.build()
builder.showFinalList()

ToExcel = ExcelWriter()
ToExcel.makeExcel(builder.getFinalUniqueList())


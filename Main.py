import ExcelPathsGetter
from SearchEngineExcel import SearchEngineExcel
from RowToObjects import RowMaker

MAINEXCELPATHNAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/DoTestow.xlsx"
EXCELPATHNAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/90408.xls"

rowObjects, rowsWithBadDKF = RowMaker.readExcelFileToSheetAndMakingObject(MAINEXCELPATHNAME)
rowObjects[0].showRowObject()
rowObjects[6].showRowObject()
print("Number of rows with bad DKF:  ", len(rowsWithBadDKF))

print("All files in dir:    ", ExcelPathsGetter.excelPaths)
print()
SearchEngineExcel.ScanFileForPatterns(EXCELPATHNAME)

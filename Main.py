import DocReader
from RowToObjects import RowMaker

EXCELPATHNAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/DoTestow.xlsx"
#DOCXPATHNAME = "D:/Poligon_Python/Faktura-VAT.docx"

#DocReader.readDoc(DOCXPATHNAME)
rowObjects, rowsWithBadDKF = RowMaker.readExcelFileToSheetAndMakingObject(EXCELPATHNAME)
rowObjects[0].showRowObject()
rowObjects[6].showRowObject()
print("Number of rows with bad DKF:  ", len(rowsWithBadDKF))



import DocReader
from ScopeForRows import RowMaker

EXCELPATHNAME = "D:/Poligon_Python/TestExcelFile.xlsx"
DOCXPATHNAME = "D:/Poligon_Python/Faktura-VAT.docx"

DocReader.readDoc(DOCXPATHNAME)

p1 = RowMaker.readExcelFileToSheetAndMakingObject(EXCELPATHNAME)
p1[0].showRowObject()



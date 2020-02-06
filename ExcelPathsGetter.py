import glob

excelPaths = []
for filename in glob.glob('//kmsrv01/OCR/EXCEL/OUTPUT/*.xls'):
    excelPaths.append(filename)

import glob

excelPaths = []
# path to mapped disk
mappedDisk = "//kmsrv01/OCR/EXCEL/OUTPUT/*.xls"
for filename in glob.glob(mappedDisk):
    excelPaths.append(filename)

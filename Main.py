from BuildFinalExcel import BuildFinalExcel
from Xlsxwriter import ExcelWriter

MAIN_EXCEL_PATH_NAME = "C:/Users/dMichalczak/TestyPyKsiegowosc/DoTestow.xlsx"
DIR_WITH_PDF_CHANGED_TO_EXCEL = "C:/Users/dMichalczak/TestyPyKsiegowosc/Pdf_to_Excel/*.xls"

builder = BuildFinalExcel(MAIN_EXCEL_PATH_NAME, DIR_WITH_PDF_CHANGED_TO_EXCEL)
builder.build()
builder.show_dkfs_wins_offices_list()

ToExcel = ExcelWriter()
ToExcel.make_excel(builder.get_final_unique_elements_list())


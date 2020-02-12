from BuildFinalExcel import BuildFinalExcel
from Xlsxwriter import ExcelWriter
    # NEONET MAIN EXCEL: "C:/Users/dMichalczak/TestyPyKsiegowosc/NEONET_Main.xlsx"
    # Dir_with_Pdf_to_excel(NEONET): "C:/Users/dMichalczak/TestyPyKsiegowosc/Pdf_to_Excel_NEONET/*.xls"
    # Path to new excel file: "C:/Users/dMichalczak/TestyPyKsiegowosc/Ksiegowosc_test_Neonet.xlsx"
    # Index of dkf column: 8
    # Index_of row_to scan: 1

    #NEO24 MAIN EXCEL: "C:/Users/dMichalczak/TestyPyKsiegowosc/NEO24_Main.xls"
    # Dir_with_Pdf_to_excel(NEO24): "C:/Users/dMichalczak/TestyPyKsiegowosc/Pdf_to_excel_NEO24/*.xls"
    # Path to new excel file: "C:/Users/dMichalczak/TestyPyKsiegowosc/Ksiegowosc_test_NEO24.xlsx"
    # Index of dkf column: 7
    # Index_of row_to scan: 2

    #NEO24 MAIN EXCEL: "C:/Users/dMichalczak/TestyPyKsiegowosc/NEO24_Main"
    # Dir_with_Pdf_to_excel(NEO24): "C:/Users/dMichalczak/TestyPyKsiegowosc/Pdf_to_excel_NEO24/*.xls"
    # Path to new excel file: "C:/Users/dMichalczak/TestyPyKsiegowosc/Ksiegowosc_test_Neonet.xlsx"

main_excel_path_name = "C:/Users/dMichalczak/TestyPyKsiegowosc/NEO24_Main.xls"
dir_with_pdf_changed_to_excel = "C:/Users/dMichalczak/TestyPyKsiegowosc/Pdf_to_excel_NEO24/*.xls"
new_excel_path = "C:/Users/dMichalczak/TestyPyKsiegowosc/Ksiegowosc_test_NEO24.xlsx"
index_of_dkf_column = 7
index_row_scan = 2

builder = BuildFinalExcel(main_excel_path_name, index_of_dkf_column, index_row_scan, dir_with_pdf_changed_to_excel)
builder.build()
builder.showFinalList()

ToExcel = ExcelWriter()
ToExcel.makeExcel(builder.getFinalUniqueList(), new_excel_path)


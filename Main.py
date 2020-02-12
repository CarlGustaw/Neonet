from BuildFinalExcel import BuildFinalExcel
from Xlsxwriter import ExcelWriter
    # NEONET MAIN EXCEL: "C:/Users/dMichalczak/TestyPyKsiegowosc/NEONET_Main.xlsx"
    # Dir_with_Pdf_to_excel(NEONET): "C:/Users/dMichalczak/TestyPyKsiegowosc/Pdf_to_Excel_NEONET/*.xls"

    #NEO24 MAIN EXCEL: "C:/Users/dMichalczak/TestyPyKsiegowosc/NEO24_Main"
    # Dir_with_Pdf_to_excel(NEO24): "C:/Users/dMichalczak/TestyPyKsiegowosc/Pdf_to_excel_NEO24/*.xls"

    #NEO24 MAIN EXCEL: "C:/Users/dMichalczak/TestyPyKsiegowosc/NEO24_Main"
    # Dir_with_Pdf_to_excel(NEO24): "C:/Users/dMichalczak/TestyPyKsiegowosc/Pdf_to_excel_NEO24/*.xls"

main_excel_path_name = "C:/Users/dMichalczak/TestyPyKsiegowosc/NEONET_Main.xlsx"
dir_with_pdf_changed_to_excel = "C:/Users/dMichalczak/TestyPyKsiegowosc/Pdf_to_Excel_NEONET/*.xls"
new_excel_path = "C:/Users/dMichalczak/TestyPyKsiegowosc/Ksiegowosc_test_Neonet.xlsx"

builder = BuildFinalExcel(main_excel_path_name, dir_with_pdf_changed_to_excel)
builder.build()
builder.showFinalList()

ToExcel = ExcelWriter()
ToExcel.makeExcel(builder.getFinalUniqueList(), new_excel_path)


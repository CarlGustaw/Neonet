# Reading an excel file using Python
import xlrd


def read_excel_file(path_name):

    # To open Workbook
    wb = xlrd.open_workbook(path_name)
    return wb.sheet_by_index(0)

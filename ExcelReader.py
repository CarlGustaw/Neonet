# Reading an excel file using Python
import xlrd


def readExcelFile(pathname):

    # To open Workbook
    wb = xlrd.open_workbook(pathname)
    sheet = wb.sheet_by_index(0)

    return sheet

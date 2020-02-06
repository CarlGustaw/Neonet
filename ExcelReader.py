import xlrd


def readExcelFile(pathname):

    # To open Workbook
    workbook = xlrd.open_workbook(pathname)
    sheet = workbook.sheet_by_index(0)

    return sheet

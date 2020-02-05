import xlsxwriter


class ExcelWriter:

    @staticmethod
    def makeExcel(listAsSheet):
        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook('KsiegowoscTest.xlsx')
        worksheet = workbook.add_worksheet()

        # Start from the first cell. Rows and columns are zero indexed.
        row = 0
        col = 0

        # Iterate over the data and write it out row by row.
        for iD_DKF, winVersion, officeVersion in listAsSheet:
            worksheet.write(row, col, iD_DKF)
            worksheet.write(row, col + 1, winVersion)
            worksheet.write(row, col + 2, officeVersion)
            row += 1

        workbook.close()

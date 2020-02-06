import xlsxwriter


class ExcelWriter:
    id_DKF = ""
    winVersion = ""
    officeVersion = ""

    def makeExcel(self, listAsSheet):
        self.id_DKF = listAsSheet[0]
        self.winVersion = listAsSheet[1]
        self.officeVersion = listAsSheet[2]
        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook('C:/Users/dMichalczak/TestyPyKsiegowosc/KsiegowoscTest.xlsx')
        worksheet = workbook.add_worksheet()

        # Start from the first cell. Rows and columns are zero indexed.
        row = 0
        col = 0

        # Iterate over the data and write it out row by row.
        worksheet.write(row, col, self.id_DKF)
        worksheet.write(row, col + 1, self.winVersion)
        worksheet.write(row, col + 2, self.officeVersion)
        row += 1

        workbook.close()
